import { Component, OnInit } from '@angular/core';
import { ExportExcelService } from '../service/export-excel.service';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.css'],
})
export class HomeComponent implements OnInit {
  formEx: any = {
    name: '',
    age: '',
    phone: '',
    email: '',
    facebook: '',
    address: '',
  };
  file?: any;
  arrayBuffer?: any;
  dataForExcel: any[] = [];

  constructor(private exportExelService: ExportExcelService) {}

  ngOnInit(): void {}

  addfile(event: any) {
    this.file = event.target.files[0];
    let fileReader = new FileReader();
    fileReader.readAsArrayBuffer(this.file);
    fileReader.onload = (e) => {
      this.arrayBuffer = fileReader.result;
      var data = new Uint8Array(this.arrayBuffer);
      var arr = new Array();
      for (var i = 0; i != data.length; ++i)
        arr[i] = String.fromCharCode(data[i]);
      var bstr = arr.join('');
      var workbook = XLSX.read(bstr, { type: 'binary' });
      var first_sheet_name = workbook.SheetNames[0];
      var worksheet = workbook.Sheets[first_sheet_name];
      worksheet['!cols'];
      var arraylist: any[] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      console.log(arraylist);
      this.formEx.name = arraylist[4][0];
      this.formEx.age = arraylist[4][1];
      this.formEx.phone = arraylist[4][2];
      this.formEx.email = arraylist[4][3];
      this.formEx.facebook = arraylist[4][4];
      this.formEx.address = arraylist[4][5];
    };
  }

  exportToExcel() {
    this.dataForExcel = Object.values(this.formEx);
    let reportData = {
      title: 'Form templates',
      data: [this.dataForExcel],
      headers: ['Name', 'Age', 'Phone Number', 'Email', 'Facebook', 'Address'],
    };
    this.exportExelService.exportExcel(reportData);
  }
}
