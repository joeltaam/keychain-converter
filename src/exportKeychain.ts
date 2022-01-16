import * as path from 'path'
import * as fs from 'fs'
import * as Excel from 'exceljs'

const currentWorkingPath = process.cwd()

type HeaderKey =
  | 'folder'
  | 'favorite'
  | 'type'
  | 'name'
  | 'notes'
  | 'fields'
  | 'reprompt'
  | 'login_uri'
  | 'login_username'
  | 'login_password'
  | 'login_totp'

const BITWARDEN_COLUMNS: Array<{ header: HeaderKey; key: HeaderKey }> = [
  { header: 'folder', key: 'folder' },
  { header: 'favorite', key: 'favorite' },
  { header: 'type', key: 'type' },
  { header: 'name', key: 'name' },
  { header: 'notes', key: 'notes' },
  { header: 'fields', key: 'fields' },
  { header: 'reprompt', key: 'reprompt' },
  { header: 'login_uri', key: 'login_uri' },
  { header: 'login_username', key: 'login_username' },
  { header: 'login_password', key: 'login_password' },
  { header: 'login_totp', key: 'login_totp' },
]

const excelHeaders = [
  'Title',
  'Login URL',
  'Login Username',
  'Login Password',
  'Additional URLs',
] as const

export const exportKeychain = (filePath: string) => {
  if (!filePath) {
    throw new Error('请输入文件地址')
  }
  const absPath = path.join(currentWorkingPath, filePath)
  if (!fs.existsSync(absPath)) {
    throw new Error('找不到文件')
  }
  const { worksheet: bwWorksheet, workbook: bwWorkbook } = bitwardenCsvCreator()
  const workbook = new Excel.Workbook()

  workbook.csv.readFile(filePath).then((worksheet) => {
    for (let i = 2; i <= worksheet.rowCount; i++) {
      const row: { [key in HeaderKey]: string } = {
        folder: '',
        favorite: '',
        type: '',
        name: '',
        notes: '',
        fields: '',
        reprompt: '',
        login_uri: '',
        login_username: '',
        login_password: '',
        login_totp: '',
      }
      excelHeaders.forEach((cell, j) => {
        const value =
          worksheet
            .getRow(i)
            .getCell(j + 1)
            .value?.toString() || ''
        switch (cell) {
          case 'Title':
            row['name'] = value
            break
          case 'Login Username':
            row['login_username'] = value
            break
          case 'Login Password':
            row['login_password'] = value
            break
          case 'Login URL':
            row['login_uri'] = value
            break
          case 'Additional URLs':
            break
        }
      })
      bwWorksheet.addRow(row)
    }
    bwWorkbook.csv
      .writeFile(path.join(currentWorkingPath, 'bitwarden_import.csv'))
      .then(() => {
        console.log('Done.')
      })
  })
}

const bitwardenCsvCreator = () => {
  const workbook = new Excel.Workbook()
  const worksheet = workbook.addWorksheet('bitwarden')
  worksheet.columns = BITWARDEN_COLUMNS
  return { workbook, worksheet }
}
