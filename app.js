var express = require('express'),
    app = express(),
    bodyParser = require('body-parser'),
    multer = require('multer'),
    xlsx = require('xlsx'),
    pdfDocument = require('pdfkit'),
    path = require('path'),
    server_port = process.env.OPENSHIFT_NODEJS_PORT || 8080,
    server_ip_address = process.env.OPENSHIFT_NODEJS_IP || '127.0.0.1',
    LABEL_EN = require('./locale/en'),
    model = require('./model/model');
app.use(bodyParser.json());

var storage = multer.diskStorage({ //multers disk storage settings
    destination: function (req, file, cb) {
        cb(null, './excelstore/')
    },
    filename: function (req, file, cb) {
        var datetimestamp = Date.now();
        cb(null, file.fieldname + '-' + datetimestamp + '.' + file.originalname.split('.')[file.originalname.split('.').length - 1])
    }
});
var upload = multer({ //multer settings
    storage: storage,
    fileFilter: function (req, file, callback) { //file filter
        /*
        if (['xls', 'xlsx'].indexOf(file.originalname.split('.')[file.originalname.split('.').length - 1]) === -1) {
            return callback(new Error('Wrong extension type'));
        }
        */
        callback(null, true);
    }
}).single('file');
/** API path that will upload the files */
app.post('/upload', function (req, res) {
    upload(req, res, function (err) {
        var filePath = req.file.path;
        console.log(path.resolve(filePath));
        if (err) {
            res.json({ statusCode: 0, description: err });
            return;
        }
        var workbook = xlsx.readFile(filePath);
        console.log(workbook.SheetNames);
        generatePDF(res);
    });
});
/*
app.get('/', function (req, res) {
    res.sendFile(__dirname + "/index.html");
});
*/
app.get('/', function (req, res) {
    renderPDF(res);
    //res.sendFile(__dirname + "/index.html");
});
app.listen(server_port, server_ip_address, function () {
    console.log('Running on ' + server_port);
});
function renderPDF(response) {
    var doc = new pdfDocument({
        layout: "portrait",
        size: "LEGAL"
    });
    doc.pipe(response);//Output is to response
    doc.registerFont('gillsansbold', './fonts/gillsansbold.ttf');
    doc.registerFont('gillsanslight', './fonts/gillsanslight.ttf')
    /*Loop should start here*/
    doc.rect(30, 1, 100, 70).fillAndStroke('black')
        .font('gillsanslight').fontSize(18).text(LABEL_EN.mainHeading, 50, 100)
        .moveTo(50, 130).lineTo(570, 130).stroke()//underline of the heading
        .font('gillsansbold').fontSize(10).text(LABEL_EN.partnerShipName, 60, 140)
        .moveTo(50, 160).lineTo(570, 160).stroke()//underline of the heading
        .font('gillsansbold').fontSize(10).text(LABEL_EN.gpStrategy, 55, 165)//table heading column 1
        .font('gillsansbold').fontSize(10).text(LABEL_EN.investmentRationale, 210, 165)//table heading column2
        .moveTo(50, 180).lineTo(570, 180).dash(1, { space: 1 }).stroke()//dashed line  of the heading
        .font('gillsanslight').fontSize(9).text(model.partnerShipName, 55, 190)//table heading
        .font('gillsanslight').fontSize(9).text(model.investRation, 210, 190, { align: 'justify', wordSpacing: 0, lineGap: 1,width:295 })
        .moveTo(50, 270).undash().lineTo(570, 270).stroke()//underline of the heading
        .font('gillsansbold').fontSize(10).text(LABEL_EN.summaryOfPartTerms, 50, 273)
        .moveTo(50, 290).lineTo(570, 290).stroke()
        .font('gillsanslight').fontSize(9).text(LABEL_EN.fundSize, 50, 295)//labels
        .font('gillsanslight').fontSize(9).text(model.fundSize, 210, 295)//value
        .font('gillsanslight').fontSize(9).text(LABEL_EN.invPeriod, 340, 295)//labels
        .font('gillsanslight').fontSize(9).text(model.investmentPeriod, 480, 295)//value
        .moveTo(50, 310).lineTo(570, 310).dash(1, { space: 1 }).stroke()//dashed line  of the heading
        .font('gillsanslight').fontSize(9).text(LABEL_EN.indFocus, 50, 315)//labels
        .font('gillsanslight').fontSize(9).text(model.industryFocus, 210, 315)//value
        .font('gillsanslight').fontSize(9).text(LABEL_EN.hurdRate, 340, 315)//labels
        .font('gillsanslight').fontSize(9).text(model.hurdleRate, 480, 315)//value
        .moveTo(50, 330).lineTo(570, 330).dash(1, { space: 1 }).stroke()//dashed line  of the heading
        .font('gillsanslight').fontSize(9).text(LABEL_EN.geoFocus, 50, 335)//labels
        .font('gillsanslight').fontSize(9).text(model.geographicFocus, 210, 335)//value
        .font('gillsanslight').fontSize(9).text(LABEL_EN.manageFees, 340, 335)//labels
        .font('gillsanslight').fontSize(9).text(model.managementFees, 480, 335)//value
        .moveTo(50, 350).lineTo(570, 350).dash(1, { space: 1 }).stroke()//dashed line  of the heading
        .font('gillsanslight').fontSize(9).text(LABEL_EN.avgTransSize, 50, 355)//labels
        .font('gillsanslight').fontSize(9).text(model.averageTransSize, 210, 355)//value
        .font('gillsanslight').fontSize(9).text(LABEL_EN.carriedIns, 340, 355)//labels
        .font('gillsanslight').fontSize(9).text(model.carriedInt, 480, 355)//value
        .moveTo(50, 370).lineTo(570, 370).dash(1, { space: 1 }).stroke()//dashed line  of the heading
        .font('gillsanslight').fontSize(9).text(LABEL_EN.partLife, 50, 375)//labels
        .font('gillsanslight').fontSize(9).text(model.partnerShipLife, 210, 375)//value
        .font('gillsanslight').fontSize(9).text(LABEL_EN.proTeamSize, 340, 375)//labels
        .font('gillsanslight').fontSize(9).text(model.professionalTeamSize, 480, 375)//value
        .moveTo(50, 390).lineTo(570, 390).dash(1, { space: 1 }).stroke()//dashed line  of the heading
        .font('gillsanslight').fontSize(9).text(LABEL_EN.offices, 340, 395)//labels
        .font('gillsanslight').fontSize(9).text(model.offices, 480, 395)//value
        /*Next Table*/
        .moveTo(50, 430).undash().lineTo(570, 430).stroke()//underline of the heading
        .font('gillsansbold').fontSize(10).text(LABEL_EN.partnerPerformance, 50, 433)
        .moveTo(50, 450).lineTo(570, 450).stroke()
        .font('gillsanslight').fontSize(9).text(LABEL_EN.lastReportNav, 50, 455)//labels
        .font('gillsanslight').fontSize(9).text(model.lastReportedNav, 210, 455)//value
        .font('gillsanslight').fontSize(9).text(LABEL_EN.nav, 340, 455)//labels
        .font('gillsanslight').fontSize(9).text(model.nav, 480, 455)//value
        .moveTo(50, 470).lineTo(570, 470).dash(1, { space: 1 }).stroke()//dashed line  of the heading
        .font('gillsanslight').fontSize(9).text(LABEL_EN.commitment, 50, 475)//labels
        .font('gillsanslight').fontSize(9).text(model.commitment, 210, 475)//value
        .font('gillsanslight').fontSize(9).text(LABEL_EN.totalIncVal, 340, 475)//labels
        .font('gillsanslight').fontSize(9).text(model.totalValuesSinceIncep, 480, 475)//value
        .moveTo(50, 490).lineTo(570, 490).dash(1, { space: 1 }).stroke()//dashed line  of the heading
    /*Loop Ends Here*/
    doc.end();

};
function generatePDF(responseObject, config, input) {
    var doc = new pdfDocument();
    doc.pipe(responseObject);
    doc.fontSize(25).text('Here is some vector graphics...', 100, 80);
    doc.scale(0.6)
        .translate(470, 130)
        .path('M 250,75 L 323,301 131,161 369,161 177,301 z')
        .fill('red', 'even-odd')
        .restore();
    doc.image('./batman.jpg', 100, 80, { width: 200, height: 100 });
    doc.end();
}