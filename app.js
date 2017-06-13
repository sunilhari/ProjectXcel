var express = require('express'),
    app = express(),
    bodyParser = require('body-parser'),
    multer = require('multer'),
    xlsx = require('xlsx'),
    pdfDocument = require('pdfkit'),
    path = require('path'),
    server_port = process.env.OPENSHIFT_NODEJS_PORT || 8080,
    server_ip_address = process.env.OPENSHIFT_NODEJS_IP || '127.0.0.1';
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
        size:"LEGAL"
    });
    doc.pipe(response);//Output is to response
    doc.registerFont('gillsansbold', './fonts/gillsansbold.ttf');
    doc.registerFont('gillsanslight', './fonts/gillsanslight.ttf')
    /*Loop should start here*/
    doc.rect(30, 1, 100, 70).fillAndStroke('black');
    doc.font('gillsanslight').fontSize(18).text('Review of Portfolio Funds.', 50, 100);
    doc.moveTo(50, 130).lineTo(570, 130).stroke();//underline of the heading
    doc.font('gillsansbold').fontSize(10).text('PartnerShip Name: ABC', 60, 140);
    doc.moveTo(50, 160).lineTo(570, 160).stroke();//underline of the heading
    doc.font('gillsansbold').fontSize(10).text('GP Strategy', 55, 165);
    doc.font('gillsansbold').fontSize(10).text('Investment Rationale', 230, 165);

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