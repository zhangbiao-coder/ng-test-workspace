import {Component, ElementRef, ViewChild} from '@angular/core';
import {ToExcelOption, ZETUtil} from 'z-export-table';

@Component({
  selector: 'app-root',
  imports: [],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss'
})
export class AppComponent {
  @ViewChild("table") table: ElementRef | undefined;

  exportsExcel() {
    let option: ToExcelOption = {
      filename: "复杂html导出Excel测试",
      userCss: `table { width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 14px; text-align: center; } th, td { border: 1px solid #ddd; padding: 10px; } .highlight { background-color: #ffeb3b; font-weight: bold; } .merged { background-color: #ffc107; } .header { background-color: #ff7043; color: white; font-weight: bold; }`
    };
    if (this.table) {
      ZETUtil.exportTableToExcel(this.table.nativeElement.outerHTML, option);
    }
  }
}
