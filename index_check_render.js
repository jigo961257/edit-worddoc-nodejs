const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const fs = require("fs");
const path = require("path");

const delay = ms => { return new Promise(resolve => setTimeout(resolve, ms)); }

let data = [
    {
        title: "title",
        filed: "filed of inovation",
        object: "object of filed"
    },
    {
        title: "title",
        filed: "filed of inovation",
        object: "object of filed"
    },
    {
        title: "title",
        filed: "filed of inovation",
        object: "object of filed"
    },

]

const final_fun = async () => {
    var make_sizzp = new PizZip();
    async function performACtion() {
        for (let i = 0; i < data.length; i++) {
            console.log("start the opration ", i);
            var content = fs
                .readFileSync(path.resolve(__dirname, 'files/test_automation.docx'), 'binary');


            // console.log(content);
            var zip = new PizZip(content, {});

            var doc = new Docxtemplater(zip, {
                delimiters: { start: "<<", end: ">>" },
            });
            // doc.loadZip(zip);
            // doc.options.delimiters.start = "<<"
            // doc.options.delimiters.end = ">>"
            // console.log(doc.setData.toString());

            function hdandel_update() {
                doc.setData({
                    title_excel_paraphrase: data[i].title + "( " + i + ")",
                    filed_of_inovation_paraphrase: data[i].filed + "( " + i + ")",
                    object_of_parapharse: data[i].object + "( " + i + ")",
                })
            }

            hdandel_update()


            try {
                // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
                doc.render()
            }
            catch (error) {
                var e = {
                    message: error.message,
                    name: error.name,
                    stack: error.stack,
                    properties: error.properties,
                }
                console.log(JSON.stringify({ error: e }));
                // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
                throw error;
            }

            var buf = doc.getZip()
                .generate({ type: 'nodebuffer' });
            let file_name = 'output_i' + i + '_.docx';
            make_sizzp.file(file_name, buf)
            fs.writeFileSync(path.resolve(__dirname, 'files/out_folder/' + file_name), buf);
            console.log("complete the opration ", i);
            await delay(5000)
        }
    }

    await performACtion();

    
    var content_1 = null;
    if (PizZip.support.uint8array) {
        content_1 = make_sizzp.generate({ type: "uint8array" });
    } else {
        content_1 = make_sizzp.generate({ type: "string" });
    }
    fs.writeFileSync(path.resolve(__dirname, 'files/out_folder/output_.zip'), content_1);
};

final_fun();

