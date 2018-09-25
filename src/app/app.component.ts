import { Component } from '@angular/core';

declare const Excel: any;

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

  color = 'green';
  sum = 0;

  onColorRows() {
    Excel.run((context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = this.color;
      return context.sync();
    });
  }

  onAddRows() {
    Excel.run((context) => {
      this.sum = 0;
      const range = context.workbook.getSelectedRange();
      range.load('values');
      return context.sync().then( () => {
        for (const rows of range.m_values) {
          for (const value of rows) {
            this.sum += +value;
          }
        }
      });
    });
  }

  private objToString (obj) {
    let str = '';
    for (const p in obj) {
      if (obj.hasOwnProperty(p)) {
        str += p + '::' + obj[p] + '\n';
      }
    }
    return str;
  }
}
