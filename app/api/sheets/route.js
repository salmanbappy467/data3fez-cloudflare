import { SignJWT, importPKCS8 } from 'jose'
import { NextResponse } from 'next/server'

export const runtime = 'edge'

// Initialize Google Auth Keys
let privateKey = process.env.GOOGLE_PRIVATE_KEY
if (privateKey) {
  privateKey = privateKey.trim()
  if (privateKey.startsWith('"') && privateKey.endsWith('"')) {
    privateKey = privateKey.substring(1, privateKey.length - 1)
  }
  if (privateKey.startsWith("'") && privateKey.endsWith("'")) {
    privateKey = privateKey.substring(1, privateKey.length - 1)
  }
  privateKey = privateKey.replace(/\\n/g, '\n').replace(/\\r/g, '')
}
const clientEmail = process.env.GOOGLE_CLIENT_EMAIL

async function getAccessToken() {
  if (!privateKey || !clientEmail) {
    throw new Error('Google credentials are not set')
  }

  const alg = 'RS256'
  const privateKeyObj = await importPKCS8(privateKey, alg)

  const jwt = await new SignJWT({
    iss: clientEmail,
    sub: clientEmail,
    aud: 'https://oauth2.googleapis.com/token',
    scope: 'https://www.googleapis.com/auth/spreadsheets'
  })
    .setProtectedHeader({ alg, typ: 'JWT' })
    .setIssuedAt()
    .setExpirationTime('1h')
    .sign(privateKeyObj)

  const response = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`
  })

  if (!response.ok) {
    const errorText = await response.text()
    throw new Error(`Failed to get access token: ${errorText}`)
  }

  const data = await response.json()
  return data.access_token
}

async function fetchGoogleSheets(path, method = 'GET', body = null) {
  const token = await getAccessToken()
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${path}`
  const reqInit = {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
  }
  if (body) {
    reqInit.body = JSON.stringify(body)
  }
  
  const response = await fetch(url, reqInit)
  if (!response.ok) {
    const errorData = await response.json().catch(() => ({ error: { message: response.statusText } }))
    throw new Error(errorData.error?.message || 'Failed to fetch Google Sheets API')
  }
  return response.json()
}

/**
 * Helper to get the correct range with escaped sheet name.
 * If sheetName is not found, it creates it.
 */
async function getEffectiveRange(spreadsheetId, sheetName, range) {
  const spreadsheet = await fetchGoogleSheets(spreadsheetId)
  let sheet = spreadsheet.sheets.find(s => s.properties.title === sheetName)

  if (!sheet) {
    // Create the sheet if it doesn't exist
    const response = await fetchGoogleSheets(`${spreadsheetId}:batchUpdate`, 'POST', {
      requests: [
        {
          addSheet: {
            properties: {
              title: sheetName,
            },
          },
        },
      ],
    })
    sheet = response.replies[0].addSheet
  }

  return `'${sheetName}'!${range}`
}

export async function POST(request) {
  try {
    const { action, sheetId, sheetName, data, row, col, value, rowIndex } = await request.json()

    if (!sheetId) throw new Error('Sheet ID is required')

    let responseData = { status: 'success' }

    switch (action) {
      case 'read': {
        const range = await getEffectiveRange(sheetId, sheetName, 'A:AC')
        const response = await fetchGoogleSheets(`${sheetId}/values/${encodeURIComponent(range)}`)
        responseData.data = response.values || []
        break
      }

      case 'add': {
        const range = await getEffectiveRange(sheetId, sheetName, 'A:A')
        await fetchGoogleSheets(`${sheetId}/values/${encodeURIComponent(range)}:append?valueInputOption=RAW`, 'POST', {
          values: [Array.isArray(data) ? data : Object.values(data)],
        })
        break
      }

      case 'update': {
        const range = await getEffectiveRange(sheetId, sheetName, `${String.fromCharCode(64 + col)}${row}`)
        await fetchGoogleSheets(`${sheetId}/values/${encodeURIComponent(range)}?valueInputOption=RAW`, 'PUT', {
          values: [[value]],
        })
        break
      }

      case 'updateRow': {
        const { row, values } = data
        const range = await getEffectiveRange(sheetId, sheetName, `A${row}:ZZ${row}`)
        await fetchGoogleSheets(`${sheetId}/values/${encodeURIComponent(range)}?valueInputOption=RAW`, 'PUT', {
          values: [values],
        })
        break
      }

      case 'batchUpdate': {
        const { updates } = data // Array of { range, values }
        await fetchGoogleSheets(`${sheetId}/values:batchUpdate`, 'POST', {
          data: updates.map(u => ({
            range: `'${sheetName}'!${u.range}`,
            values: u.values
          })),
          valueInputOption: 'RAW',
        })
        break
      }

      case 'delete': {
        const index = rowIndex || row
        if (!index) throw new Error('Row index is required for deletion')

        const spreadsheet = await fetchGoogleSheets(sheetId)
        let sheet = spreadsheet.sheets.find(s => s.properties.title === sheetName)

        if (!sheet) {
          // If sheet doesn't exist, nothing to delete from it
          break
        }

        const gid = sheet.properties.sheetId

        await fetchGoogleSheets(`${sheetId}:batchUpdate`, 'POST', {
          requests: [
            {
              deleteDimension: {
                range: {
                  sheetId: gid,
                  dimension: 'ROWS',
                  startIndex: index - 1,
                  endIndex: index,
                },
              },
            },
          ],
        })
        break
      }

      default:
        throw new Error('Unsupported action')
    }

    return NextResponse.json(responseData)
  } catch (error) {
    console.error('Google Sheets API Error:', error)
    return NextResponse.json(
      { status: 'error', message: error.message },
      { status: 500 }
    )
  }
}
