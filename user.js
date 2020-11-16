var fs = require('fs');
var openXml = require('openxml');
var path = require('path');
var PizZip = require('pizzip');
var Docxtemplater = require('docxtemplater');
var officejs = require('@microsoft/office-js');

class TxtStyle
{
    constructor(fontName,fontSize,fontSealing,paragraphIndent,isMultiline,cellWidth) {
        this.fontName = fontName;
        this.fontSize = fontSize;
        this.fontSealing = fontSealing;
        this.paragraphIndent = paragraphIndent;
        this.isMultiline = isMultiline;
        this.cellWidth = cellWidth;
    }
}

class DocumentFields
{
    constructor(name,value, tableValue, style,
                tableValueMaxLen = null, valueMaxLen = 0, type = 0,
                rowHeight = null) {
        this.name = name;
        if (valueMaxLen === 0 || value.length <= valueMaxLen)
            this.value = value;
        this.tableValue = tableValue;
        this.style = style;
        this.tableValueMaxLen = tableValueMaxLen;
        this.valueMaxLen = valueMaxLen;
        this.type = type;
        this.rowHeight = rowHeight
    }
    constructor() {
        this.name = "Times New Roman";
        this.type = 0;
        this.valueMaxLen = 0;
        this.value = "";
        this.tableValueMaxLen = null;
        this.tableValue = null;
        this.style = new TxtStyle(
            "Times New Roman",
            10,
            0,
            0,
            false,
            0
        )
    }

    calcRowHeight(){
        let height = 0.0;
        this.rowHeight.forEach(el => height += el);
        return height;
    }

    calcRowHeight(index){
        let height = 0.0;
        this.rowHeight.forEach(el => height += el);
        return height;
    }

    Copy()
    {
        return JSON.parse(JSON.stringify(this));
    }

}

class ConvertToPDF{
    constructor(documentFields,templatePath, PDFPath, docSavePath, fileName) {
        this.documentFields = documentFields;
        this.templatePath = templatePath;
        this.PDFPath = PDFPath;
        this.docSavePath = docSavePath;
        this.fileName = fileName;
    }

    FillPDF()
    {
        var dat = new Date();
        dat = dat.getUTCMilliseconds();
        let name = this.fileName + "_" + dat;
        let fileNameDoc = this.docSavePath + name + ".docx";
        let fileNamePDF = this.PDFPath + "Doc_" + dat + ".pdf";
//         fs.createReadStream(this.templatePath,fileNameDoc);
// //Load the docx file as a binary
//         var content = fs
//             .readFileSync(path.resolve(__dirname, 'test.docx'), 'binary');
//
//         var zip = new PizZip(content);
//         var doc;
//         try {
//             doc = new Docxtemplater(zip);
//         } catch(error) {
//             // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
//             errorHandler(error);
//         }
        Word.run(function (context) {

            var doc = context.document;

            //get the bookmark range object by its name
            var bookmarkRange=doc.getBookmarkRangeOrNullObject("cscasenumber01");

            //insert a data and replace thee bookmark range
            bookmarkRange.insertText("test data",Word.InsertLocation.replace);

            // Synchronize the document state by executing the queued commands,
            return context.sync();

        }).catch(errorHandler);

        this.documentFields.forEach((el) =>{
            switch (el.type) {
                case 1:
                {
                    var bookmarkRange=doc.getBookmarkRangeOrNullObject("cscasenumber01");
                    bookmarkRange.load();

                    return context.sync()
                        .then(function() {
                            if (bookmarkRange.isNullObject) {
                                // handle case of null object here
                            } else {
                                bookmarkRange.insertText("test data",Word.InsertLocation.replace);
                            }
                        })
                        .then(context.sync)
                    //let Name = el.name;
                    //doc.setData({
                    //    Name: el.value
                    //});

                    break;
                }
                case 2:
                {

                    break;
                }
                case 3:
                {

                    break;
                }
                default:
                {
                    console.log("Ошибочка, таких ошибок нет");
                }
            }
        },this)

        try {
            // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
            doc.render()
        }
        catch (error) {
            // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
            errorHandler(error);
        }

        var buf = doc.getZip()
            .generate({type: 'nodebuffer'});

// buf is a nodejs buffer, you can either write it to a file or do anything else with it.
        fs.writeFileSync(path.resolve(__dirname, 'output.docx'), buf);
    }
}



function User(name, age){

    this.name = name;
    this.age = age;
    this.displayInfo = function(){

        console.log(`Имя: ${this.name}  Возраст: ${this.age}`);
    }
}
User.prototype.sayHi = function() {
    console.log(`Привет, меня зовут ${this.name}`);
};

module.exports = User;