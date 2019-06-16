import { Component, OnInit, NgZone} from '@angular/core';
import * as XLSX from 'xlsx';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
  constructor(private ngZone: NgZone) {
    this.zone = ngZone;
  }
  zone: any;
  title = 'Angular 8.x';
  listOfData = [];
  columnHeaders = [];
  ngOnInit(): void {
    console.log('init');
  }
  changeFileHandler(e: any) {
    const files = e.target.files;
    const fileReader = new FileReader();
    let workbook: any;
    fileReader.onload = (ev) => {
            try {
                const data = (ev.target as FileReader).result;
                workbook = XLSX.read(data, {
                    type: 'binary'
                }); // 以二进制流方式读取得到整份excel表格对象
                this.listOfData = []; // 存储获取到的数据
            } catch (e) {
                console.log('文件类型不正确');
                return;
            }
            // 表格的表格范围，可用于判断表头是否数量是否正确
            let fromTo = '';
            // 遍历每张表读取
            for (const sheet in workbook.Sheets) {
                if (workbook.Sheets.hasOwnProperty(sheet)) {
                    fromTo = workbook.Sheets[sheet]['!ref'];
                    console.log(fromTo);
                    this.listOfData = this.listOfData.concat(XLSX.utils.sheet_to_json(workbook.Sheets[sheet]));
                    break; // 如果只取第一张表，就取消注释这行
                }
            }
            // 在控制台打印出来表格中的数据
            const firstRow = this.listOfData[0];
            console.log(Object.keys(firstRow));
            console.log(this.listOfData);
            this.zone.run(() => {
              this.columnHeaders = Object.keys(firstRow);
            });
        };
        // 以二进制方式打开文件
    fileReader.readAsBinaryString(files[0]);
    }
}
