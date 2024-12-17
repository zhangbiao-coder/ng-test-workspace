import {Component, ElementRef, ViewChild} from '@angular/core';
import {ExcelUtils} from './util/ExcelUtils';

@Component({
  selector: 'app-root',
  imports: [],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss'
})
export class AppComponent {
  @ViewChild("table") table:ElementRef|undefined;

  exportsExcel() {
    if(this.table){
      ExcelUtils.exportTableToExcel(this.table.nativeElement.outerHTML,{
        filename:"复杂html导出Excel测试",
        userCss:`table { width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 14px; text-align: center; } th, td { border: 1px solid #ddd; padding: 10px; } .highlight { background-color: #ffeb3b; font-weight: bold; } .merged { background-color: #ffc107; } .header { background-color: #ff7043; color: white; font-weight: bold; }`
      })
    }
  }
}
