import { Component } from '@angular/core';

import * as XLSX from 'xlsx';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';

type AOA = any[][];

@Component({
  selector: 'app-sheet',
  templateUrl: './sheet.component.html',
  styleUrls: ['./sheet.component.css']
})
export class SheetJSComponent {
  colValue: any;
  data: AOA = [];
  actualData: AOA = [];
  wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
  fileName: string = 'SheetJS.xlsx';

  onFileChange(evt: any) {
    /* wire up file reader */
    const target: DataTransfer = <DataTransfer>evt.target;
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      /* read workbook */
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });

      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      /* save data */
      this.data = <AOA>XLSX.utils.sheet_to_json(ws, { header: 1 });
      this.actualData = <AOA>XLSX.utils.sheet_to_json(ws, { header: 1 });
    };
    reader.readAsBinaryString(target.files[0]);
  }

  // You can add a trackBy to determine if the list must or must not be     reloaded. The code below seems to solve the issue:
  customTrackBy(index: number, obj: any): any {
    return index;
  }

  save(): void {
    let isValidData = true;
    console.log(this.data);
    if (this.data.length > 0) {
      for (let r = 1; r < this.data.length; r++) {
        for (let c = 0; c < this.data[r].length; c++) {
          if (
            this.data[0][c] === 'EmployeeName' &&
            this.data[r][c] !== this.getProperCase(this.data[r][c])
          ) {
            window.alert('Employee name must be in proper case.');
            isValidData = false;
            break;
          }

          if (this.data[0][c] === 'Age' && this.data[r][c] < 25) {
            window.alert('Age must be greater than or equal to 25.');
            isValidData = false;
            break;
          }
        }
      }
    }

    if (isValidData) {
      this.exportExcel();
    }
  }

  getProperCase(str): string {
    return str.replace(/\w\S*/g, function(txt) {
      return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    });
  }

  exportExcel() {
    const workBook = new Workbook();
    const workSheet = workBook.addWorksheet('test');

    for (let r = 0; r < this.data.length; r++) {
      const row = workSheet.addRow([
        this.data[r][0],
        this.data[r][1],
        this.data[r][2]
      ]);
      for (let c = 0; c < this.data[r].length; c++) {
        // color code modified cell
        if (this.actualData[r][c] !== this.data[r][c]) {
          const col = row.getCell(c + 1);
          col.font = { color: { argb: '#DC1A1A' } };
        }
      }
    }

    // export excel
    workBook.xlsx.writeBuffer().then(data => {
      let blob = new Blob([data], {
        type:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      fs.saveAs(blob, 'Exported.xlsx');
    });
  }
}
