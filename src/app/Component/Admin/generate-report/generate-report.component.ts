import { Component, ElementRef, OnInit, ViewChild } from '@angular/core';
import { Router } from '@angular/router';
import { Questions } from 'src/app/Model/Questions';
import { UserService } from 'src/app/Service/user.service';
import * as XLSX from 'xlsx';  

@Component({
  selector: 'app-generate-report',
  templateUrl: './generate-report.component.html',
  styleUrls: ['./generate-report.component.css']
})
export class GenerateReportComponent implements OnInit {
  @ViewChild('TABLE', { static: false }) TABLE: ElementRef;  
  title = 'Excel';  
  ExportTOExcel() {  
    const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet(this.accounts);  
    const wb: XLSX.WorkBook = XLSX.utils.book_new();  
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');  
    XLSX.writeFile(wb, 'Sheet.xls');  
  }  
  ExportTOCsv() {  
    const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet(this.accounts);  
    const wb: XLSX.WorkBook = XLSX.utils.book_new();  
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');  
    XLSX.writeFile(wb, 'Sheet.csv');  
  }  

  ExportTOText() {  
    const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet(this.accounts);  
    const wb: XLSX.WorkBook = XLSX.utils.book_new();  
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');  
    XLSX.writeFile(wb, 'Sheet.txt');  
  }  
  ExportTOPdf() {  
    const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet(this.accounts);  
    const wb: XLSX.WorkBook = XLSX.utils.book_new();  
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');  
    XLSX.writeFile(wb, 'Sheet.html');  
  }  

  accounts: Questions[];
  userName:string;

  constructor(private router: Router, private userService: UserService) {} 

  ngOnInit() {
    if (localStorage.getItem('username') != null) {
      this.userService.getPolicy().subscribe((data) => {
        this.accounts = data;
      });
    } else this.router.navigate(['/login']);
  }
 
  Search() {
    if(this.userName == "")
    {
      this.ngOnInit();
    }
    else
    {
      this.accounts = this.accounts.filter( res => {
        return res.userName.toLocaleLowerCase().match(this.userName.toLocaleLowerCase());
      });
    }
  }
  key:string = 'userName';
  reverse:boolean = false;
  sort(key: string)
  {
    this.key = key;
    this.reverse = !this.reverse;
  }
}
