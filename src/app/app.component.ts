import { Component, OnInit } from '@angular/core';
import * as GC from '@grapecity/spread-sheets';
import * as ExcelIO from "@grapecity/spread-excelio";
import '@grapecity/spread-sheets-charts';
import { jsonData } from './app.data';
import { ExcelUtility } from './ExcelUtility';
import { saveAs } from "file-saver"; // npm package: "file-saver": "^1.3.8"
import { Workbook } from 'igniteui-angular-excel';
import { WorkbookFormat } from 'igniteui-angular-excel';
import { WorkbookSaveOptions } from 'igniteui-angular-excel';
declare var saveAs: any;
import * as XLSX from 'xlsx';
(window as any).GC = GC;
type AOA = any[][];

@Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit {

    public columnDefs = [
        { field: "make" },
        { field: "model" },
        { field: "price" }
    ];

    // specify the data
    public rowData = [
        { make: "Toyota", model: "Celica", price: 35000 },
        { make: "Ford", model: "Mondeo", price: 32000 },
        { make: "Porsche", model: "Boxter", price: 72000 }
    ];

    // let the grid know which columns and what data to use
    public gridOptions = {
        columnDefs: null,
        rowData: null
    };

    spread: GC.Spread.Sheets.Workbook;
    importExcelFile: any;
    exportFileName = "export.xlsx";
    password: string;
    data: AOA = [];
    constructor() {
    }

    ngOnInit(): void {
        // setup the grid after the page has finished loading
        // document.addEventListener('DOMContentLoaded', function () {
        //     var gridDiv = document.querySelector('#myGrid');
        //     new agGrid.Grid(gridDiv, gridOptions);
        // });
    }


    initSpread($event: any) {
        // debugger
        this.spread = $event.spread;
        let spread = this.spread;
        spread.options.calcOnDemand = true;
        spread.fromJSON(jsonData);
    }
    changeFileDemo(e: any) {
        debugger
        this.importExcelFile = e.target.files[0];
        this.uploadExcel(this.importExcelFile);
    }
    changePassword(e: any) {
        this.password = e.target.value;
    }
    changeExportFileName(e: any) {
        this.exportFileName = e.target.value;
    }
    changeIncremental(e: any) {
        document.getElementById('loading-container').style.display = e.target.checked ? "block" : "none";
    }
    loadExcel(e: any) {
        debugger
        let spread = this.spread;
        let excelIo = new ExcelIO.IO();
        let excelFile = this.importExcelFile;
        this.uploadExcel(excelFile);
        let password = this.password;
        let incrementalEle = document.getElementById("incremental") as HTMLInputElement;
        let loadingStatus = document.getElementById("loadingStatus") as HTMLInputElement;

        // here is excel IO API
        excelIo.open(excelFile, function (json: any) {
            let workbookObj = json;
            if (incrementalEle.checked) {
                spread.fromJSON(workbookObj, {
                    incrementalLoading: {
                        loading: function (progress: any) {
                            progress = progress * 100;
                            loadingStatus.value = progress;
                        },
                        loaded: function () {
                        }
                    }
                });
            } else {
            }
        }, function (e: any) {
            // process error
            alert(e.errorMessage);
        }, { password: password });
    }

    uploadExcel(evt: any) {

        /* wire up file reader */
        const target: any = <DataTransfer>(evt);
        // if (target.files.length !== 1) throw new Error('Cannot use multiple files');
        const reader: FileReader = new FileReader();
        reader.onload = (e: any) => {
            /* read workbook */
            const bstr: string = e.target.result;
            const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });

            /* grab first sheet */
            const wsname: string = wb.SheetNames[0];
            const ws: XLSX.WorkSheet = wb.Sheets[wsname];

            /* save data */
            debugger
            this.data = <AOA>(XLSX.utils.sheet_to_json(ws, { header: 1 }));

            // Map to remove empty array
            this.data = this.data.filter(ele => ele.length);
            //   console.log("this.data this.data" , this.data);

            //   // var ws = wb.Sheets[wb.SheetNames[0]];
            //   var html_string = XLSX.utils.sheet_to_html(ws, { id: "xslxTable", editable: true });
            //   console.log("html_string html_string" , html_string);
            //   document.getElementById("container").innerHTML = html_string;

        };
        reader.readAsBinaryString(target);

        // try {

        //   const fileName = e.target.files[0].name;

        //   import('xlsx').then(xlsx => {
        //     let workBook = null;
        //     let jsonData = null;
        //     const reader = new FileReader();
        //     // const file = ev.target.files[0];
        //     reader.onload = (event) => {
        //       const data = reader.result;
        //       workBook = xlsx.read(data, { type: 'binary' });
        //       jsonData = workBook.SheetNames.reduce((initial, name) => {
        //         const sheet = workBook.Sheets[name];
        //         initial[name] = xlsx.utils.sheet_to_json(sheet);
        //         return initial;
        //       }, {});
        //       this.products = jsonData[Object.keys(jsonData)[0]];
        //       console.log(this.products);

        //     };
        //     reader.readAsBinaryString(e.target.files[0]);
        //   });

        // } catch (e) {
        //   console.log('error', e);
        // }

    }


    saveExcel(e: any) {
        debugger
        let spread = this.spread;
        let excelIo = new ExcelIO.IO();

        let fileName = this.exportFileName;
        let password = this.password;
        if (fileName.substr(-5, 5) !== '.xlsx') {
            fileName += '.xlsx';
        }

        let json = spread.toJSON();

        // here is excel IO API
        excelIo.save(json, function (blob: any) {
            saveAs(blob, fileName);
        }, function (e: any) {
            // process error
            console.log(e);
        }, { password: password });

    }

    public static save(workbook: Workbook, fileNameWithoutExtension: string): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            const opt = new WorkbookSaveOptions();
            opt.type = "blob";

            workbook.save(opt, (d) => {
                const fileExt = ExcelUtility.getExtension(workbook.currentFormat);
                const fileName = fileNameWithoutExtension + fileExt;
                saveAs(d as Blob, fileName);
                resolve(fileName);
            }, (e) => {
                reject(e);
            });
        });
    }



    // saveExcelSheet() {
    //     debugger
    //     // var elt = document.getElementsByClassName('gc-floatingobject-background-cover gc-no-user-select gcdv-flexdv');
    //     var wb = XLSX.utils.json_to_sheet(this.data);
    //     XLSX.writeFile(wb, "Test Hamada");
    // }
}