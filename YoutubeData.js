fs = require('fs');
const YoutubeDlWrap = require("youtube-dl-wrap");
const youtubeDlWrap = new YoutubeDlWrap("./node_modules/youtube-dl/bin/youtube-dl.exe");
var excel = require('excel4node');
var workbook = new excel.Workbook();
var worksheet = workbook.addWorksheet('Sheet 1');
const xlsx = require("xlsx");
const spreadsheet = xlsx.readFile('./links.xlsx');
const sheets = spreadsheet.SheetNames;
const firstSheet = spreadsheet.Sheets[sheets[0]]; //sheet 1 is index 0


(async () => {

    let links = []
    let tracks = []
    let artists = []

    for (let i = 1; ; i++) {
        const firstColumn = firstSheet['A' + i];
        if (!firstColumn) {
            break;
        }
        links.push(firstColumn.h);
    }

    for (let index = 0; index < links.length; index++) {
        try {
            console.log(`Fetch ${links[index]}`)
            let metadata = await youtubeDlWrap.getVideoInfo(links[index]);
            if (metadata.track == undefined || metadata.track == null) {
                tracks.push(`error`)
                artists.push(`error`)
            }
            else {
                tracks.push(metadata.track)
                artists.push(metadata.artist)
            }

        } catch (error) {
            console.log(error)
            tracks.push(`error`)
            artists.push(`error`)
        }
    }
    const outputFields = [
        "Link",
        "Track",
        "Artist"
    ]

    for (let i = 0; i < outputFields.length; i++) {
        worksheet.cell(1, i + 1).string(outputFields[i])
    }

    for (let index = 0; index < links.length; index++) {
        worksheet.cell(index + 2, 1).string(links[index])
        worksheet.cell(index + 2, 2).string(tracks[index])
        worksheet.cell(index + 2, 3).string(artists[index])
    }
    workbook.write('Results.xlsx')
    console.log('Done!')
})()
