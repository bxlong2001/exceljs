const ExcelJS = require('exceljs');
const fs = require('fs')

const handleExcel = async () => {
    const data = fs.readFileSync('data.json')
    const jsonData = JSON.parse(data)
    const workbook = new ExcelJS.Workbook();
    const excelFile = await workbook.xlsx.readFile('soKho.xlsx');
    const worksheet = excelFile.getWorksheet("06-03-2023")
    let sum = 0
    let lastRow = 0
    jsonData.forEach((tp, index) => {
        worksheet.getCell(`A${index+11}`).value = index+1
        worksheet.getCell(`B${index+11}`).value = tp.tenThucPham
        worksheet.getCell(`C${index+11}`).value = tp.donViTinh
        worksheet.getCell(`D${index+11}`).value = tp.soLuong
        worksheet.getCell(`E${index+11}`).value = tp.donGia
        worksheet.getCell(`F${index+11}`).value = tp.soLuong * tp.donGia
        sum += worksheet.getCell(`F${index+11}`).value
        lastRow = index+12
    })
    worksheet.getCell(`B${lastRow}`).value = 'Tổng cộng'
    worksheet.getCell(`F${lastRow}`).value = sum
    await workbook.xlsx.writeFile('xuat.xlsx')
    console.log('File saved');
}

handleExcel()