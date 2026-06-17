const ExcelJS = require('exceljs');

class DataConverter {
  static toJSON(data, pretty = true) {
    return pretty ? JSON.stringify(data, null, 2) : JSON.stringify(data);
  }

  static toCSV(data, fields, includeBOM = true) {
    if (!Array.isArray(data) || data.length === 0) {
      return includeBOM ? '\uFEFF' : '';
    }

    const headers = fields.map(f => f.label || f.name);
    const fieldNames = fields.map(f => f.name);

    let csv = headers.join(',') + '\n';

    data.forEach(row => {
      const values = fieldNames.map(name => {
        let value = row[name];
        if (value === null || value === undefined) {
          return '';
        }
        if (typeof value === 'string') {
          value = value.replace(/"/g, '""');
          if (value.includes(',') || value.includes('\n') || value.includes('"')) {
            value = `"${value}"`;
          }
        }
        return value;
      });
      csv += values.join(',') + '\n';
    });

    return includeBOM ? '\uFEFF' + csv : csv;
  }

  static async toExcel(data, fields) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Mock Data Generator';
    workbook.created = new Date();

    const worksheet = workbook.addWorksheet('数据');

    const headers = fields.map(f => f.label || f.name);
    const fieldNames = fields.map(f => f.name);

    const headerRow = worksheet.addRow(headers);

    headerRow.eachCell((cell) => {
      cell.font = {
        bold: true,
        color: { argb: 'FFFFFFFF' },
        size: 12
      };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF4472C4' }
      };
      cell.alignment = {
        horizontal: 'center',
        vertical: 'middle'
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFD9D9D9' } },
        left: { style: 'thin', color: { argb: 'FFD9D9D9' } },
        bottom: { style: 'thin', color: { argb: 'FFD9D9D9' } },
        right: { style: 'thin', color: { argb: 'FFD9D9D9' } }
      };
    });

    headerRow.height = 28;

    if (Array.isArray(data) && data.length > 0) {
      data.forEach(row => {
        const values = fieldNames.map(name => {
          const value = row[name];
          if (value === null || value === undefined) {
            return '';
          }
          return this.formatValue(value);
        });
        const dataRow = worksheet.addRow(values);
        dataRow.eachCell((cell, colNumber) => {
          cell.alignment = {
            vertical: 'middle'
          };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFEFEFEF' } },
            left: { style: 'thin', color: { argb: 'FFEFEFEF' } },
            bottom: { style: 'thin', color: { argb: 'FFEFEFEF' } },
            right: { style: 'thin', color: { argb: 'FFEFEFEF' } }
          };
          const fieldType = fields[colNumber - 1] ? fields[colNumber - 1].type : null;
          if (fieldType === 'number') {
            cell.alignment = { horizontal: 'right', vertical: 'middle' };
          }
        });
      });
    }

    const MAX_COL_WIDTH = 50;
    const MIN_COL_WIDTH = 10;
    worksheet.columns.forEach((column, i) => {
      let maxLength = headers[i] ? headers[i].length : MIN_COL_WIDTH;
      if (data && data.length > 0) {
        const sampleSize = Math.min(data.length, 100);
        for (let r = 0; r < sampleSize; r++) {
          const value = data[r][fieldNames[i]];
          if (value !== null && value !== undefined) {
            const cellLength = String(value).length;
            if (cellLength > maxLength) {
              maxLength = cellLength;
            }
          }
        }
      }
      const width = Math.min(Math.max(maxLength + 4, MIN_COL_WIDTH), MAX_COL_WIDTH);
      column.width = width;
    });

    worksheet.views = [
      { state: 'frozen', ySplit: 1 }
    ];

    worksheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: headers.length }
    };

    return workbook.xlsx.writeBuffer();
  }

  static downloadJSON(res, data, filename = 'data.json') {
    const jsonData = this.toJSON(data);
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(jsonData);
  }

  static downloadCSV(res, data, fields, filename = 'data.csv') {
    const csvData = this.toCSV(data, fields);
    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(csvData);
  }

  static async downloadExcel(res, data, fields, filename = 'data.xlsx') {
    const buffer = await this.toExcel(data, fields);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(Buffer.from(buffer));
  }

  static formatValue(value) {
    if (value === null || value === undefined) {
      return 'null';
    }
    if (typeof value === 'boolean') {
      return value ? '是' : '否';
    }
    if (typeof value === 'number') {
      return value.toLocaleString();
    }
    return String(value);
  }
}

module.exports = DataConverter;
