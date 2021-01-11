import { Component, ElementRef,  ViewChild } from "@angular/core";

import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import PizZipUtils from "pizzip/utils/index.js";
import { saveAs } from "file-saver";

function loadFile(url, callback) {
  PizZipUtils.getBinaryContent(url, callback);
}

@Component({
  selector: "app-product-list",
  templateUrl: "./product-list.component.html",
  styleUrls: ["./product-list.component.css"]
})
export class ProductListComponent {
  @ViewChild("file",null) inputView:ElementRef;
  data={
    "PropNum":"CD0000.01",
    "Date":"May 14, 2020",
    "ProjectName":"Dime Box ISD 2020 Campus Improvements",
    "Price":"5.00",
    "Purchaser":"Americo Perez",
    "inclusions":[
        {"projectName":"testName1",
        "projectText":"textTest1",
        "projectText2":"text2Test1",
        "grandTotal":"5,000.00",
        "inclusion":"1"},
        {"projectName":"testName2",
        "projectText":"textTest2",
        "projectText2":"text2Test2",
        "grandTotal":"6,000.00",
        "inclusion":"1"},
        {"inclusion":"1"},
    ],
    "exclusions":[
        {"exclusion":"1"},
        {"exclusion":"1"},
        {"exclusion":"1"},
        {"exclusion":"1"},
        {"exclusion":"1"},
        {"exclusion":"1"},
        {"exclusion":"1"},
        {"exclusion":"1"},
        {"exclusion":"1"},
        {"exclusion":"1"}
    ]
}
  content:any;
  generate() {
    
      var zip = new PizZip(this.content);
      var doc = new Docxtemplater(zip);
      doc.setData(this.data);
      try {
        // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
        doc.render();
      } catch (error) {
        // The error thrown here contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).
        function replaceErrors(key, value) {
          if (value instanceof Error) {
            return Object.getOwnPropertyNames(value).reduce(function(
              error,
              key
            ) {
              error[key] = value[key];
              return error;
            },
            {});
          }
          return value;
        }
        console.log(JSON.stringify({ error: error }, replaceErrors));

        if (error.properties && error.properties.errors instanceof Array) {
          const errorMessages = error.properties.errors
            .map(function(error) {
              return error.properties.explanation;
            })
            .join("\n");
          console.log("errorMessages", errorMessages);
          // errorMessages is a humanly readable message looking like this :
          // 'The tag beginning with "foobar" is unopened'
        }
        throw error;
      }
      var out = doc.getZip().generate({
        type: "blob",
        mimeType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      }); //Output the document using Data-URI
      saveAs(out, "output.docx");
  }

  share() {
    window.alert("The product has been shared!");
  }

  onClick($event){
    this.inputView.nativeElement.click()
  }

  FileEmitter(files:FileList){
    var file = files.item(0);
    var reader = new FileReader();
    reader.onload=()=>{
      this.content = reader.result;
      // console.log(this.content);
    }
    reader.readAsBinaryString(file);
  }
}

/*
Copyright Google LLC. All Rights Reserved.
Use of this source code is governed by an MIT-style license that
can be found in the LICENSE file at http://angular.io/license
*/
