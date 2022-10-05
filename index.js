const HtmlDocx = require('html-docx-js');
const fs = require('fs');

const generateMarkup = (document, {  orientation  = '', border = {}, header = {}, footer = {}, pageNumberAlignment, pageNumberPosition }) =>{

    let htmlContent = `<html>`;
    // Starting Head Tags
    htmlContent += `<head>`;

    // Default Style of the document
    htmlContent += `<style type="text/css"  media="print">table { border-collapse: collapse; } table, td, th { border: 0px solid black; }</style>`;

    // Page Section Settings
    htmlContent += `<style type="text/css"  media="print"> 
            @page Section1 {
                margin:0in 0in 0in 0in;
                mso-header-margin:0.5in;
                mso-header: h1;
                mso-footer-margin:0.5in;
                mso-footer: f1;
                mso-paper-source:0;
            }
            div.Section1 {page:Section1;}
            div.MsoNormal{
                mso-style-parent:"";
                margin-top : ${border.top};
                margin-bottom: ${border.bottom};
                margin-left: ${border.left};
                margin-right: ${border.right};
                word-spacing: 0;
                font-family:"Arial";
                mso-fareast-font-family:"Arial";
            }
            pre, li, div, p, span, form, h1, h2, h3, h4, h5, h6, table, thead, th, tbody, tr, td, img, input, textarea, dd, dt, dl{
                margin:0in;
                padding : 0in
                word-spacing: 0;
            }
            ol, ul {
                margin: 0 !important;
                word-spacing: 0 !important;
            }
            p.headerFooter { margin:0in; text-align: center; }

            table#hrdftrtbl
            {
                margin:0in 0in 0in 900in;
                width:1px;
                height:1px;
                overflow:hidden;
                visibility: hidden;
            }
            table, tr, td{
                border : 0px solid #FFF;
            }
            p.MsoFooter, li.MsoFooter, div.MsoFooter {
                mso-pagination:widow-orphan;
                tab-stops:center 3.0in right 6.0in;
                text-align : center;
            }
            p.MsoHeader, li.MsoHeader, div.MsoHeader {
                mso-pagination:widow-orphan;
                tab-stops:center 3.0in right 6.0in;
                text-align : center;
            }
            </style>`;

    // Ending Head Tags
    htmlContent += `</head>`;


    // Start Body Tags
    htmlContent += `<body>`;
    // Start Page Section
    htmlContent += `<div class=Section1 style="margin:1in 1in 1in 1in">`;
    // Table
    htmlContent += `<table id='hrdftrtbl' border='1' cellspacing='0' cellpadding='0' style='margin-left:50in;visibility: hidden;'>`;
    htmlContent += `<tr style='height:1pt;mso-height-rule:exactly;visibility: hidden;'>`;
    htmlContent += `<td>`;
    if(header?.contents?.default){
        htmlContent += `<div style='mso-element:header' id=h1>`;
            htmlContent += `<p class="MsoHeader">`;
                htmlContent += `<table border="0" width= "100%" cellpadding="0" cellspacing="0" style="border-bottom:0px solid #736D6E">`;
                    htmlContent += `<tr>`;
                        htmlContent += `${header?.contents?.default}`;
                        if(pageNumberPosition === 'header'){
                            htmlContent += `<td style="text-align:${pageNumberAlignment}"> - <span style='mso-field-code:" PAGE "'></span> - </td>`;
                        }
                    htmlContent += `</tr>`;
                htmlContent += `</table>`;
            htmlContent += `</p>`;
        htmlContent += '</div>';
    }
    htmlContent += `</td>`;
    htmlContent += `<td>`;
    if(footer?.contents?.default){
        htmlContent += `<div style='mso-element:footer' id=f1>`;
        htmlContent += `<p class="MsoFooter">`;   
            htmlContent += `<table style="border-top: 0px solid black;" width="100%" cellpadding="0" cellspacing="0">`; 
                htmlContent += `<tr>`;
                    if(pageNumberPosition === 'footer'){
                        htmlContent += `<td style="text-align:${pageNumberAlignment}"> - <span style='mso-field-code:" PAGE "'></span> - </td>`;
                    }
                    htmlContent += `${footer?.contents?.default}`;
                htmlContent += `</tr>`;
            htmlContent += `</table>`; 
        htmlContent += `</p>`;
        htmlContent += `</div>`;
    }
    htmlContent += `</td>`;
    htmlContent += `</tr>`;
    htmlContent += '</table>';
    htmlContent += '<div class=MsoNormal>';
    if(document.html){
        htmlContent += `${document.html}`;
    }
    htmlContent += '</div>';
    htmlContent += `</div>`;
    htmlContent += `</body>`;


    htmlContent += `</html>`;

    return htmlContent;



}

const pageSize = async(pageFormat) =>{
    return new Promise((resolve, reject) => {
        let pageSize = '';
        switch(pageFormat){
            case "Letter":
                pageSize = '21.59cm 27.94cm';
                break;
            case "Tabloid":
                pageSize = '27.94cm 43.18cm';
                break;
            case "Legal":
                pageSize = '21.59cm 35.56cm';
                break;
            case "Statement":
                pageSize = '13.97cm 21.59cm';
                break;
            case "Executive":
                pageSize = '18.41cm 26.67cm';
                break;
            case "A3":
                pageSize = '29.7cm 42cm';
                break;
            case "A4":
                pageSize = '21cm 29.7cm';
                break;
            case "A5":
                pageSize = '14.8cm 21cm';
                break;

        }
        resolve(pageSize);
    })
}


const createDoc = async (markupText, {format, orientation, border, header, footer, margins, pageNumberAlignment, pageNumberPosition}) => {
    return new Promise(async(resolve, reject) => {
        const inputMarkup = generateMarkup(markupText, { orientation, border, header, footer, pageNumberAlignment, pageNumberPosition })
        var outputFileName = `${markupText.path}`;
        var docx = HtmlDocx.asBlob(inputMarkup, { orientation, margins : { ...margins }, size : format});
        fs.writeFile(outputFileName, docx, function (err) {
            if (err) {
                reject(err)
                return 0
            };
            resolve(`${outputFileName}`)
        });
    })
}

module.exports = {
    createDoc
}