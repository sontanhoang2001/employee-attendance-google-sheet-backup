//const momentFake = require('moment');
require('dotenv').config();

const { GoogleSpreadsheet } = require('google-spreadsheet');
const moment = require("moment-timezone");
const QRCode = require('qrcode');
const ExcelJS = require('exceljs');
const removeDiacritics = require('remove-diacritics');
const { createCanvas, loadImage } = require('canvas');


const PRIVATE_KEY = process.env.PRIVATE_KEY;
const CLIENT_EMAIL = process.env.CLIENT_EMAIL;
const SHEET_ID = process.env.SHEET_ID;
// var SHEET_ID = '';


let getListUsers = async (req, res) => {
    try {
        // Initialize the sheet - doc ID is the long id in the sheets URL
        const doc = new GoogleSpreadsheet(SHEET_ID);

        // Initialize Auth - see more available options at https://theoephraim.github.io/node-google-spreadsheet/#/getting-started/authentication
        await doc.useServiceAccountAuth({
            client_email: CLIENT_EMAIL,
            private_key: PRIVATE_KEY,
        });

        await doc.loadInfo(); // loads document properties and worksheets

        const sheet = doc.sheetsByTitle["Chấm công"]; // or use doc.sheetsById[id] or doc.sheetsByTitle[title]
        const rows = await sheet.getRows();
        var listUser = rows.map(user => user["Họ Tên (Nhận Từ Form)"]);
        listUser = listUser.filter((item, index) => {
            return listUser.indexOf(item) === index;
        });
        // console.log("rows: ", listUser)
        res.json({ listUser });
    } catch (error) {
        res.json({ listUser: [] });
    }
};

let getLoginpage = async (req, res) => {
    // SHEET_ID = req.query.id;
    // if (SHEET_ID !== '') {
    // console.log("SHEET_ID: ", SHEET_ID)
    const listUser = ["Sơn Tấn Hoàng", "Nguyễn Hữu Ái"];
    return res.render("loginpage.ejs", { loginStatus: '', code: req.query.code, dataInputed: { phone: '', fullName: '' } })
    // }
    // res.send("Phiên làm việc đã hết hạn. Vui lòng quét lại mãi QR code.");
};
let getTimekeepingPage = async (req, res) => {
    return res.render("timekeepingPage.ejs")
};

let login = async (req, res) => {
    try {
        res.setHeader('Content-Type', 'text/html');

        // if (SHEET_ID === '') {
        //     res.send("Phiên làm việc đã hết hạn. Vui lòng quét lại mãi QR code.");
        // }
        const fullNameInput = req.body.fullName;
        const phone = req.body.phone;

        if (phone == '') {
            const errorMessage = 'Vui lòng nhập số điện thoại!';
            res.render("loginpage.ejs", { loginStatus: errorMessage, code: req.query.code, dataInputed: { phone: phone, fullName: fullNameInput } });
            return;
        }

        const regex = /^(?:\+84|0)(?:1\d{9}|3\d{8}|5\d{8}|7\d{8}|8\d{8}|9\d{8})$/;
        if (!regex.test(phone)) {
            const errorMessage = 'Số điện thoại không họp lệ. Vui lòng nhập lại!';
            res.render("loginpage.ejs", { loginStatus: errorMessage, code: req.query.code, dataInputed: { phone: phone, fullName: fullNameInput } });
            return;
        }


        const checkin = req.body.checkin;
        const checkout = req.body.checkout;

        // Initialize the sheet - doc ID is the long id in the sheets URL
        const doc = new GoogleSpreadsheet(SHEET_ID);

        // Initialize Auth - see more available options at https://theoephraim.github.io/node-google-spreadsheet/#/getting-started/authentication
        await doc.useServiceAccountAuth({
            client_email: CLIENT_EMAIL,
            private_key: PRIVATE_KEY,
        });

        await doc.loadInfo(); // loads document properties and worksheets

        const sheet = doc.sheetsByTitle["Danh sách nhân viên"]; // or use doc.sheetsById[id] or doc.sheetsByTitle[title]

        const rows = await sheet.getRows();
        const userInfor = rows.find(row => row["Số Điện Thoại"] === phone);
        // console.log("userInfor: ", userInfor);

        // user không có trong danh sách
        if (!userInfor) {
            if (fullNameInput == '') {
                const errorMessage = 'Số điện không tồn tại trong hệ thống. Vui lòng nhập thêm "Họ và tên" để tiếp tục điểm danh!';
                res.render("loginpage.ejs", { loginStatus: errorMessage, code: req.query.code, dataInputed: { phone: '', fullName: '' } });
                return;
            }

            if (req.body.doForOther == "1") {
                const errorMessage = 'Tài khoản của bạn chưa có danh sách không được phép điểm danh hộ!';
                res.render("loginpage.ejs", { loginStatus: errorMessage, code: req.query.code, dataInputed: { phone: phone, fullName: '' } });
                return;
            }
        } else {
            // nếu có trong ds, nhưng điểm danh hộ nhưng chưa check
            // đã nhập "Họ và tên" và chưa check vào "Điểm danh hộ"
            if (fullNameInput != "" && req.body.doForOther != "1") {
                // Nếu tên nhập và khác tên userInfor thì thông bao
                if (userInfor["Họ Tên"] != fullNameInput) {
                    const errorMessage = 'Vui lòng check vào ô điểm danh hộ để xác nhận và tiếp tục!';
                    res.render("loginpage.ejs", { loginStatus: errorMessage, code: req.query.code, dataInputed: { phone: phone, fullName: '' } });
                    return;
                }
            }
        }

        // check (check-in) or (check-out)
        if (checkin) {
            // console.log("check-in");
            insertGoogleSheet(req, res, userInfor);
        } else if (checkout) {
            // console.log("check-out");
            updateGoogleSheet(req, res, userInfor);
        }

    }
    catch (e) {
        const errorMessage = 'Đã có lỗi xảy ra!';
        res.render("loginpage.ejs", { loginStatus: errorMessage, code: req.query.code, dataInputed: { phone: phone, fullName: fullNameInput } });
    }
}

// Tạo một đối tượng Moment đại diện cho thời điểm hiện tại
// var currentDate = momentFake('2023-02-22T23:50:00');

let insertGoogleSheet = async (req, res, userInfor) => {
    try {
        // Đặt múi giờ cho server
        moment.tz.setDefault("Asia/Ho_Chi_Minh");

        // Lấy thời gian hiện tại theo múi giờ đã đặt
        let currentDate = moment().tz("Asia/Ho_Chi_Minh");

        // let currentDate = new Date();

        const format = "DD/MM/YYYY HH:mm:ss";

        let formatedDate = moment(currentDate).format(format);

        // Initialize the sheet - doc ID is the long id in the sheets URL
        const doc = new GoogleSpreadsheet(SHEET_ID);

        // Initialize Auth - see more available options at https://theoephraim.github.io/node-google-spreadsheet/#/getting-started/authentication
        await doc.useServiceAccountAuth({
            client_email: CLIENT_EMAIL,
            private_key: PRIVATE_KEY,
        });

        await doc.loadInfo(); // loads document properties and worksheets

        const sheet = doc.sheetsByTitle["Chấm công"]; // or use doc.sheetsById[id] or doc.sheetsByTitle[title]
        // const rows = await sheet.getRows();

        const fullNameInput = req.body.fullName;
        // false: chưa có account
        let phone = `${req.body.phone}`,
            fullNameForm = req.body.fullName,
            fullName = '',
            toDate = moment(currentDate).format("DD/MM/YYYY"),
            timeStart = formatedDate,
            workCode = req.query.code;

        // Đảm vảo nhập "Họ và tên"
        if (fullNameInput != "") {
            if (req.body.doForOther == "1") {
                if (userInfor["Họ Tên"] == fullNameInput) {
                    const errorMessage = 'Không thể tự điểm danh hộ cho chính bạn!';
                    res.render("loginpage.ejs", { loginStatus: errorMessage, code: req.query.code, dataInputed: { phone: phone, fullName: '' } });
                    return;
                } else {
                    phone = `(Điểm danh hộ) - ${phone}`;

                }
            }
        }

        // true: đã có account
        if (userInfor) {
            // nếu người dùng không nhập tên sẽ tự động điền tên vào 
            if (fullNameInput == '') {
                fullNameForm = userInfor["Họ Tên"];
            } else {
                fullNameForm = fullNameInput;
            }
            fullName = userInfor["Họ Tên"];
        }


        // Kiểm tra điểm danh trong ngày
        // const rowToCheck = rows.find(row => row["Số Điện Thoại"] === req.body.phone && row["Ngày"] === toDate && row["Ca Làm Việc"] === getTimeSlot(currentDate));
        // if (rowToCheck) {
        //     const errorMessage = `Bạn đã điểm danh (check-in) "${rowToCheck["Ca Làm Việc"]}" rồi. Không được phép thực hiện nữa!`;
        //     res.render("loginpage.ejs", { loginStatus: errorMessage, id: SHEET_ID });
        // } else {
        // Làm tròn thời gian
        // timeStart = roundTimeWorking(currentDate);

        await sheet.addRow(
            {
                "Số Điện Thoại": `'${phone}`,
                "Họ Tên (Nhận Từ Form)": fullNameForm,
                "Họ Tên (Danh Sách Đã Lưu)": fullName,
                "Ngày": toDate,
                // "Ca Làm Việc": getTimeSlot(currentDate),
                "Điểm Danh Lần Đầu": timeStart,
                "Rời Khỏi Lần Cuối": "Chưa",
                "Mã Tiệc": workCode
            });

        const message = `Điểm danh (check-in) thành công!`;
        res.render("timekeepingPage.ejs", { checkStatus: message, phone: phone, fullName: fullName, fullNameForm: fullNameForm, timeStart: formatedDate, timeEnd: '', code: req.query.code });
        // }
    }
    catch (e) {
        const errorMessage = `Điểm danh (check-in) thất bại!`;
        res.render("loginpage.ejs", { loginStatus: errorMessage, code: req.query.code, dataInputed: { phone: phone, fullName: fullNameInput } });
    }
}

let updateGoogleSheet = async (req, res, userInfor) => {
    try {
        // Đặt múi giờ cho server
        moment.tz.setDefault("Asia/Ho_Chi_Minh");

        // Lấy thời gian hiện tại theo múi giờ đã đặt
        let currentDate = moment().tz("Asia/Ho_Chi_Minh");

        // let currentDate = new Date();

        const format = "DD/MM/YYYY HH:mm:ss";

        let formatedDate = moment(currentDate).format(format);
        var formatedDay = moment(currentDate).format("DD/MM/YYYY");

        const doc = new GoogleSpreadsheet(SHEET_ID);

        await doc.useServiceAccountAuth({
            client_email: CLIENT_EMAIL,
            private_key: PRIVATE_KEY,
        });

        await doc.loadInfo();
        const sheet = doc.sheetsByTitle['Chấm công'];

        const rows = await sheet.getRows();

        // đầu vào
        // số đt
        // so sánh ngày hiện tại

        const fullNameInput = req.body.fullName;
        let phone = req.body.phone,
            fullNameForm = fullNameInput,
            fullName = '',
            timeStart = '';


        // Đảm vảo nhập "Họ và tên"
        if (fullNameInput != "") {
            if (req.body.doForOther == "1") {
                if (userInfor["Họ Tên"] == fullNameInput) {
                    const errorMessage = 'Không thể tự điểm danh hộ cho chính bạn!';
                    res.render("loginpage.ejs", { loginStatus: errorMessage, code: req.query.code, dataInputed: { phone: phone, fullName: '' } });
                    return;
                } else {
                    phone = `(Điểm danh hộ) - ${phone}`;

                }
            }
        }

        // true: đã có account
        if (userInfor) {
            if (fullNameInput == '') {
                fullNameForm = userInfor["Họ Tên"];
            } else {
                fullNameForm = fullNameInput;
            }
            fullName = userInfor["Họ Tên"];
        }

        // console.log("getTimeSlot(currentDate): ", getTimeSlot(currentDate))
        // console.log("Ca Làm Việc ", rowToCheck["Ca Làm Việc"])

        // const rowToUpdate = rows.find(row => row["Số Điện Thoại"] === phone && row["Ngày"] === formatedDay && row["Rời Khỏi Lần Cuối"] === "Chưa");
        // console.log("rowToUpdate: ", rowToUpdate)

        // nếu điểm danh dùm, thì khỏi check số điện thoại
        // console.log("filteredRows:", filteredRows)

        const filteredRows = rows.filter(row => row['Số Điện Thoại'] === phone && row["Họ Tên (Nhận Từ Form)"] === fullNameForm && row['Ngày'] === formatedDay);
        // console.log("filteredRows:", filteredRows)
        const rowToUpdate = filteredRows[filteredRows.length - 1]; // lấy row mới nhất
        // console.log(rowToUpdate); // in ra row mới nhất tìm được

        if (rowToUpdate["Rời Khỏi Lần Cuối"] === "Chưa") {
            timeStart = rowToUpdate["Điểm Danh Lần Đầu"];
            rowToUpdate["Số Điện Thoại"] = `'${phone}`;
            rowToUpdate["Điểm Danh Lần Đầu"] = rowToUpdate["Điểm Danh Lần Đầu"];
            rowToUpdate["Rời Khỏi Lần Cuối"] = formatedDate;
            await rowToUpdate.save();

            // return res.send("Updating data in Google Sheet succeeds!");
            const message = `Điểm danh (check-out) thành công!`;
            res.render("timekeepingPage.ejs", { checkStatus: message, phone: phone, fullName: fullName, fullNameForm: fullNameForm, timeStart: timeStart, timeEnd: formatedDate, code: req.query.code });
        } else {
            const errorMessage = `Rời đi (check-out) thất bại!`;
            res.render("loginpage.ejs", { loginStatus: errorMessage, code: req.query.code, dataInputed: { phone: phone, fullName: '' } });
        }

    } catch (e) {
        const errorMessage = `Điểm danh (check-out) thất bại!`;
        res.render("loginpage.ejs", { loginStatus: errorMessage, code: req.query.code, dataInputed: { phone: phone, fullName: '' } });
    }
};

let getQRpage = async (req, res) => {
    return res.render("qrPage.ejs", { qr: '', code: '' });
};

// let createQR = async (req, res) => {
//     // var currentUrl = 'Vui lòng nhập địa chỉ url từ google sheet!';
//     // const url = req.body.url;

//     // if (url != '') {
//     //     const sheetId = url.match(/[-\w]{25,}/)[0];
//     //     // console.log(sheetId); // in ra 1nwBi8vdQO1E_8WXFbBlVB3CBiLptDaQ-jT0fg-V1Tr4
//     //     currentUrl = req.protocol + '://' + req.get('host') + "/?id=" + sheetId;
//     //     // console.log(currentUrl);
//     // }

//     const code = req.body.code;
//     // const currentUrl = req.protocol + '://' + req.get('host') + "/chamcong/?code=" + code;
//     const currentUrl = "https://cuoihoidangkhoa.com.vn/chamcong/?code=" + code;

//     let qr = await QRCode.toDataURL(currentUrl);
//     // return res.send(img);
//     return res.render("qrPage.ejs", { qr: qr })
// };


let createQR = async (req, res) => {
    const code = req.body.code;
    const currentUrl = "/chamcong/?code=" + code;
    const maxLength = 30;

    let lines = [];
    let currentLine = '';

    let words = `Mã tiệc: ${code}`;
    // Tách chuỗi vào các dòng tối đa ? ký tự
    for (let i = 0; i < words.length; i++) {
        let word = words[i];
        let tempLine = currentLine + word;

        if (tempLine.trim().length > maxLength) {
            lines.push(currentLine.trim());
            currentLine = word;
        } else {
            currentLine = tempLine;
        }
    }

    lines.push(currentLine.trim());


    let qr = await QRCode.toDataURL(currentUrl);

    const canvas = createCanvas(300 + words.length, 300 + words.length);
    const context = canvas.getContext('2d');

    const img = await loadImage(qr);
    // Tính toán tọa độ x và y cho vẽ hình ảnh vào giữa canvas
    var x = (canvas.width - img.width) / 2;
    // var y = (canvas.height - img.height) / 2;

    // Vẽ hình ảnh vào giữa canvas
    context.drawImage(img, x, 0);
    context.font = 'bold 16px Arial';
    context.fillStyle = '#000';
    // context.fillText(`Mã tiệc: ${code}`, x, 210);

    // Xuống dòng khi dòng vượt quá chiều rộng của canvas
    let y = 210;
    for (let line of lines) {
        context.fillText(line, x - words.length / 2, y + words.length / 2);
        y += 25;
    }

    //context.textAlign = 'center';

    const finalQr = canvas.toDataURL('image/png');

    return res.render("qrPage.ejs", { qr: finalQr, code: code })
};



let getSalaryPage = async (req, res) => {
    const listRecord = [];
    const type = req.params.type;

    const unitPrice = 0;
    switch (type) {
        case 'date': {
            return res.render("salaryByDate.ejs", {
                errorMessage: '',
                listRecord: listRecord, totalSalary: '', dataInputed: {
                    startDate: '', endDate: '', unitPrice: unitPrice
                }
            })
            break;
        }
        case 'code': {
            return res.render("salaryByCode.ejs", {
                errorMessage: '',
                listRecord: listRecord, totalSalary: '', dataInputed: {
                    eventCode: '', unitPrice: unitPrice
                }
            })
            break;
        }
        case 'detail': {
            return res.render("salaryByDetail.ejs", { errorMessage: '', listRecord: listRecord, totalTime: '', totalHours: '', totalSalary: 0, unitPrice: 0, dataInputed: {} })
            break;
        }
        default:
            res.redirect('/chamcong/report/date');
    }
};

var salaryDataExport = [];
var salaryDataInputedExport = [];

let salary = async (req, res) => {
    // Initialize the sheet - doc ID is the long id in the sheets URL
    const doc = new GoogleSpreadsheet(SHEET_ID);

    // Initialize Auth - see more available options at https://theoephraim.github.io/node-google-spreadsheet/#/getting-started/authentication
    await doc.useServiceAccountAuth({
        client_email: CLIENT_EMAIL,
        private_key: PRIVATE_KEY,
    });

    await doc.loadInfo(); // loads document properties and worksheets

    const sheet = doc.sheetsByTitle["Chấm công"]; // or use doc.sheetsById[id] or doc.sheetsByTitle[title]
    const rows = await sheet.getRows();

    const formatDate = "DD/MM/YYYY";

    const fullName = req.body.fullName;
    const phone = req.body.phone;
    const eventCode = req.body.eventCode;
    const startDate = req.body.startDate;
    const endDate = req.body.endDate;
    let unitPrice = req.body.unitPrice;
    var listRecord = [];

    const type = req.params.type;
    switch (type) {
        case 'date': {
            if (startDate == '' || endDate == '') {
                return res.render("salaryByDate.ejs", {
                    errorMessage: 'Bạn chưa nhập "Từ Ngày" "Đến Ngày"!',
                    listRecord: listRecord, totalSalary: '', dataInputed: {
                        startDate: startDate, endDate: endDate, unitPrice: unitPrice
                    }
                });
            }

            if (unitPrice <= 0 || unitPrice == '') {
                return res.render("salaryByDate.ejs", {
                    errorMessage: 'Nhập "Đơn Giá" phải lớn hơn 0, "Đơn Giá" không được bỏ trống!',
                    listRecord: listRecord, totalSalary: totalSalary, dataInputed: {
                        startDate: startDate, endDate: endDate, unitPrice: unitPrice
                    }
                });
            }


            // If an event code is provided, filter the rows to include only employees who worked at the event
            listRecord = rows.filter(row => row["Đã Xuất Báo Cáo"] !== "X" && moment(row["Ngày"], formatDate).isBetween(moment(startDate), moment(endDate), null, '[]'));
            // console.log("Nhân viên: ", groupBy(listRecord, "Họ Tên (Nhận Từ Form)"));

            // tính paymant
            let totalSalary = 0;
            listRecord.forEach(row => {
                let doItFor = (row['Họ Tên (Nhận Từ Form)'] === row['Họ Tên (Danh Sách Đã Lưu)']) ? '' : row['Họ Tên (Danh Sách Đã Lưu)'];
                row["Người Chấm Công Hộ"] = doItFor;
                let money = parseInt(row["Tổng Giờ"]) * parseInt(unitPrice);
                row["Thành Tiền"] = currencyFormat(money);
                row["Thành Tiền Excel"] = money;
                row["Tổng Giờ"] = formatHourMinutes(parseFloat(row["Tổng Giờ"]));
                totalSalary += money;
                row["Đơn Giá"] = currencyFormat(unitPrice);
                row["Đơn Giá Excel"] = parseInt(unitPrice);
            });

            salaryDataExport = listRecord;
            salaryDataInputedExport = {
                totalSalary: totalSalary, dataInputed: {
                    eventCode: eventCode, startDate: moment(startDate).format(formatDate), endDate: moment(endDate).format(formatDate)
                }
            };

            return res.render("salaryByDate.ejs", {
                errorMessage: '',
                listRecord: listRecord, totalSalary: currencyFormat(totalSalary), dataInputed: {
                    startDate: startDate, endDate: endDate, unitPrice: unitPrice
                }
            });
            break;
        }
        case 'code': {
            if (eventCode == '') {
                return res.render("salaryByCode.ejs", {
                    errorMessage: 'Bạn chưa nhập mã tiệc!',
                    listRecord: listRecord, totalSalary: '', dataInputed: {
                        eventCode: '', unitPrice: unitPrice
                    }
                });
            }

            if (unitPrice <= 0 || unitPrice == '') {
                return res.render("salaryByCode.ejs", {
                    errorMessage: 'Nhập "Đơn Giá" phải lớn hơn 0, "Đơn Giá" không được bỏ trống!',
                    listRecord: listRecord, totalSalary: totalSalary, dataInputed: {
                        eventCode: eventCode, unitPrice: unitPrice
                    }
                });
            }

            if (eventCode) {
                // If an event code is provided, filter the rows to include only employees who worked at the event
                listRecord = rows.filter(row => row["Đã Xuất Báo Cáo"] !== "X" && row["Mã Tiệc"] === eventCode);
                // console.log("Nhân viên: ", listRecord);
            }

            // tính paymant
            let totalSalary = 0;
            listRecord.forEach(row => {
                let doItFor = (row['Họ Tên (Nhận Từ Form)'] === row['Họ Tên (Danh Sách Đã Lưu)']) ? '' : row['Họ Tên (Danh Sách Đã Lưu)'];
                row["Người Chấm Công Hộ"] = doItFor;
                let money = parseInt(row["Tổng Giờ"]) * parseInt(unitPrice);
                row["Thành Tiền"] = currencyFormat(money);
                row["Thành Tiền Excel"] = money;
                row["Tổng Giờ"] = formatHourMinutes(parseFloat(row["Tổng Giờ"]));
                totalSalary += money;
                row["Đơn Giá"] = currencyFormat(unitPrice);
                row["Đơn Giá Excel"] = parseInt(unitPrice);
            });

            // const listRecord = rows.find(row => row["Số Điện Thoại"] === phone);

            salaryDataExport = listRecord;
            salaryDataInputedExport = {
                totalSalary: totalSalary, dataInputed: {
                    eventCode: eventCode
                }
            };

            return res.render("salaryByCode.ejs", {
                errorMessage: '',
                listRecord: listRecord, totalSalary: currencyFormat(totalSalary), dataInputed: {
                    eventCode: eventCode, unitPrice: unitPrice
                }
            });
            break;
        }
        case 'detail': {
            if (fullName == '' && phone == '') {
                return res.render("salaryByDetail.ejs", {
                    errorMessage: 'Vui lòng nhập "Họ Tên" hoặc "Số Điện Thoại"!',
                    listRecord: listRecord, totalTime: '', totalHours: '', totalSalary: '', unitPrice: 0, dataInputed: {
                        fullName: fullName, phone: phone, startDate: startDate, endDate: endDate, unitPrice: unitPrice
                    }
                });
            }

            if (startDate == '' || endDate == '') {
                return res.render("salaryByDetail.ejs", {
                    errorMessage: 'Bạn chưa nhập "Từ Ngày" "Đến Ngày"!',
                    listRecord: listRecord, totalTime: '', totalHours: '', totalSalary: '', unitPrice: 0, dataInputed: {
                        fullName: fullName, phone: phone, startDate: startDate, endDate: endDate, unitPrice: unitPrice
                    }
                });
            }

            if (unitPrice <= 0 || unitPrice == '') {
                return res.render("salaryByDetail.ejs", {
                    errorMessage: 'Nhập "Đơn Giá" phải lớn hơn 0, "Đơn Giá" không được bỏ trống!',
                    listRecord: listRecord, totalTime: '', totalHours: '', totalSalary: '', unitPrice: 0, dataInputed: {
                        fullName: fullName, phone: phone, startDate: startDate, endDate: endDate, unitPrice: unitPrice
                    }
                });
            }

            // If an event code is provided, filter the rows to include only employees who worked at the event
            listRecord = rows.filter(row => row["Đã Xuất Báo Cáo"] !== "X" && moment(row["Ngày"], formatDate).isBetween(moment(startDate), moment(endDate), null, '[]') && (row["Họ Tên (Nhận Từ Form)"] === fullName || row["Số Điện Thoại"] === phone));
            // console.log("Nhân viên: ", listRecord);

            // const listRecord = rows.find(row => row["Số Điện Thoại"] === phone);
            let totalHours = 0;
            listRecord.forEach(row => {
                totalHours += parseFloat(convertCommaToDot(row["Tổng Giờ"]));
                let doItFor = (row['Họ Tên (Nhận Từ Form)'] === row['Họ Tên (Danh Sách Đã Lưu)']) ? '' : row['Họ Tên (Danh Sách Đã Lưu)'];
                row["Người Chấm Công Hộ"] = doItFor;
                row["totalTime"] = formatHourMinutes(parseFloat(convertCommaToDot(row["Tổng Giờ"])));
            });
            let totalTime = formatHourMinutes(parseFloat(totalHours));
            let totalSalary = parseInt(totalHours * unitPrice);

            salaryDataExport = listRecord;
            salaryDataInputedExport = {
                totalTime: totalTime, totalSalary: totalSalary, unitPrice: parseInt(unitPrice), dataInputed: {
                    fullName: fullName, phone: phone, startDate: startDate, endDate: endDate
                }
            };
            return res.render("salaryByDetail.ejs", {
                errorMessage: '',
                listRecord: listRecord, totalTime: totalTime, totalHours: totalHours, totalSalary: currencyFormat(totalSalary), unitPrice: currencyFormat(unitPrice), dataInputed: {
                    fullName: fullName, phone: phone, startDate: startDate, endDate: endDate, unitPrice: unitPrice
                }
            });
            break;
        }
        default:
    }
};

let exportExcel = (req, res) => {
    const type = req.params.type;

    // lấy ngày tháng hiện tại
    moment.tz.setDefault("Asia/Ho_Chi_Minh"); // Đặt múi giờ mặc định là Asia/Ho_Chi_Minh
    const now = moment(); // Lấy thời gian hiện tại theo múi giờ đã đặt mặc định
    const day = now.date(); // Lấy ngày hiện tại
    const month = now.month() + 1; // Lấy tháng hiện tại (chú ý phải cộng thêm 1 vì tháng bắt đầu từ 0)
    const year = now.year(); // Lấy năm hiện tại

    let fileName = '';
    let arrayFile = [`Bang-luong-${salaryDataInputedExport.dataInputed.startDate} - ${salaryDataInputedExport.dataInputed.endDate}`,
        'Ma-tiec', 'Chi-tiet-bang-luong'];
    let typeIndex = 0;
    if (type == 'date') {
        typeIndex = 0;
        fileName = arrayFile[0];
    } else if (type == 'code') {
        typeIndex = 1;
        fileName = `${arrayFile[1]} - ${salaryDataInputedExport.dataInputed.eventCode}`;
    } else if (type == 'detail') {
        typeIndex = 2;
        if (salaryDataInputedExport.dataInputed.fullName) {
            // bỏ dấu tiếng việt
            fileName = `${arrayFile[2]} - ${removeDiacritics(salaryDataInputedExport.dataInputed.fullName)}`;
        }
        if (salaryDataInputedExport.dataInputed.phone) {
            fileName = `${arrayFile[2]} - ${removeDiacritics(salaryDataInputedExport.dataInputed.phone)}`;
        }
    }

    // Tạo một Workbook mới
    const workbook = new ExcelJS.Workbook();

    // Tạo một Worksheet mới từ mảng JSON
    const worksheetArray = ["THOIVU-Time", "THOIVU-Tiec", "CHITIET"];
    const worksheet = workbook.addWorksheet(worksheetArray[typeIndex]);

    // Thêm dòng tiêu đề
    // const titleRow = worksheet.addRow(["BẢNG CHI TIẾT TIỀN LƯƠNG TIỀN CÔNG"]);
    // titleRow.font = { size: 16, bold: true };
    // worksheet.addRow([]);

    const nameCty = worksheet.addRow(["CÔNG TY TNHH MTV DỊCH VỤ HÔN LỄ ĐĂNG KHOA"]);
    nameCty.font = { bold: true, size: 14, name: 'Times New Roman' };
    const addressCty = worksheet.addRow(["183 Võ Văn Kiệt, An Thới, Bình Thủy, Cần Thơ"]);
    addressCty.font = { bold: true, size: 14, name: 'Times New Roman' };
    worksheet.addRow([]);

    // vị trí của tiêu đề
    let localtionTitleSheet = "A4:E4";
    if (type == 'date' || type == 'code') {
        localtionTitleSheet = "A4:F4";
    }

    let titleSheetArray = ["BẢNG LƯƠNG NHÂN VIÊN THỜI VỤ - THỐNG KÊ THEO TIME", "BẢNG LƯƠNG NHÂN VIÊN THỜI VỤ - THỐNG KÊ THEO TIỆC", "BẢNG CHI TIẾT TIỀN LƯƠNG TIỀN CÔNG"];
    let titleSheet = worksheet.addRow([titleSheetArray[typeIndex]]);
    worksheet.mergeCells(localtionTitleSheet);
    titleSheet.font = { bold: true, size: 16, name: 'Times New Roman' };
    titleSheet.alignment = { horizontal: "center" };

    if (type == 'date') {
        const dateRow = worksheet.addRow([`Từ ngày: ${salaryDataInputedExport.dataInputed.startDate} đến ngày: ${salaryDataInputedExport.dataInputed.endDate}`]);
        worksheet.mergeCells("A5:F5");
        dateRow.font = { bold: true, size: 14, name: 'Times New Roman' };
        dateRow.alignment = { horizontal: "center" };
        worksheet.addRow(["", "", "", "", "", ""]);

    } else if (type == 'code') {
        const dateRow = worksheet.addRow([`Nơi làm việc: ${salaryDataInputedExport.dataInputed.eventCode}`]);
        worksheet.mergeCells("A5:F5");
        dateRow.font = { bold: true, size: 14, name: 'Times New Roman' };
        dateRow.alignment = { horizontal: "center" };
        worksheet.addRow(["", "", "", "", "", ""]);

    } else if (type == 'detail') {
        const nameRow = worksheet.addRow([`Họ tên NV: ${(salaryDataInputedExport.dataInputed.fullName == '') ? '...........' : salaryDataInputedExport.dataInputed.fullName} - Số phone: ${(salaryDataInputedExport.dataInputed.phone == '') ? '........' : salaryDataInputedExport.dataInputed.phone}`]);
        worksheet.mergeCells("A5:E5");
        nameRow.font = { bold: true, size: 14, name: 'Times New Roman' };
        nameRow.alignment = { horizontal: "center" };

        const dateRow = worksheet.addRow([`Từ ngày: ${salaryDataInputedExport.dataInputed.startDate} đến ngày: ${salaryDataInputedExport.dataInputed.endDate}`]);
        worksheet.mergeCells("A6:E6");
        dateRow.font = { bold: true, size: 14, name: 'Times New Roman' };
        dateRow.alignment = { horizontal: "center" };
    }

    worksheet.addRow([]);

    if (type == 'date' || type == 'code') {
        // header
        const headerRow = worksheet.addRow(["ĐIỆN THOẠI", "HỌ TÊN NHÂN VIÊN", "HỌ TÊN NGƯỜI CHẤM CÔNG HỘ", "GIỜ CÔNG", "ĐƠN GIÁ (đồng/giờ)", "THÀNH TIỀN (đồng)"]);
        headerRow.font = { bold: true, size: 14, name: 'Times New Roman' };
        headerRow.alignment = { horizontal: "center" };

        // Set column widths
        worksheet.getColumn("A").width = 40;
        worksheet.getColumn("B").width = 30;
        worksheet.getColumn("C").width = 50;
        worksheet.getColumn("D").width = 20;
        worksheet.getColumn("E").width = 30;
        worksheet.getColumn("F").width = 30;

        // Thêm dữ liệu cho các dòng
        salaryDataExport.forEach(item => {
            const salaryData = worksheet.addRow([item["Số Điện Thoại"], item["Họ Tên (Nhận Từ Form)"], item["Người Chấm Công Hộ"], item["Tổng Giờ"], item["Đơn Giá Excel"], item["Thành Tiền Excel"]]);
            salaryData.font = { size: 14, name: 'Times New Roman' };
            salaryData.alignment = { horizontal: "center" };
        });
        worksheet.addRow([]);

        // Get the range of cells
        const nextRow = salaryDataExport.length;
        let start = { row: 8, col: 1 },
            end = { row: nextRow, col: 6 };
        // console.log("nextRow: ", nextRow);

        const totalHour = worksheet.addRow(["Tổng Cộng", "", "", "", "", salaryDataInputedExport.totalSalary]);
        totalHour.font = { bold: true, size: 14, name: 'Times New Roman' };
        totalHour.alignment = { horizontal: "center" };
        worksheet.mergeCells(`A${end.row + start.row + 2}:E${end.row + start.row + 2}`);
        worksheet.addRow([]);

        const dateCreate = worksheet.addRow(["", "", "", `Cần Thơ, ngày ${day} tháng ${month} năm ${year}`]);
        dateCreate.font = { bold: true, size: 14, name: 'Times New Roman', italic: true };
        dateCreate.alignment = { horizontal: "center" };
        worksheet.mergeCells(`D${end.row + start.row + 4}:F${end.row + start.row + 4}`);

        const signature = worksheet.addRow(["", "", "", `LẬP BIỂU`]);
        signature.font = { bold: true, size: 14, name: 'Times New Roman' };
        signature.alignment = { horizontal: "center" };
        worksheet.mergeCells(`D${end.row + start.row + 5}:F${end.row + start.row + 5}`);

        // start.row = 8;
        // end.row = length.record;
        // start.col = 1;
        // end.col = 6;
        for (let i = start.row; i <= end.row + start.row + 2; i++) {
            // tạo border cho các cell
            for (let j = start.col; j <= end.col; j++) {
                worksheet.getCell(i, j).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            }

            // Format cell A1 as currency
            // bắt đầu từ col = 5, col kết thúc = 6;
            for (let j = 5; j <= end.col; j++) {
                const cell = worksheet.getCell(i, j);
                cell.numFmt = '#,##0 ₫';
            }
        }
    } else if (type == 'detail') {
        // header
        const headerRow = worksheet.addRow(["HỌ TÊN NGƯỜI CHẤM CÔNG HỘ", "TIME BẮT ĐẦU", "TIME KẾT THÚC", "GIỜ CÔNG", "ĐỊA CHỈ LÀM VIỆC"]);
        headerRow.font = { bold: true, size: 14, name: 'Times New Roman' };
        headerRow.alignment = { horizontal: "center" };

        // Set column widths
        worksheet.getColumn("A").width = 50;
        worksheet.getColumn("B").width = 30;
        worksheet.getColumn("C").width = 30;
        worksheet.getColumn("D").width = 20;
        worksheet.getColumn("E").width = 50;

        // Thêm dữ liệu cho các dòng
        salaryDataExport.forEach(item => {
            const salaryData = worksheet.addRow([item["Người Chấm Công Hộ"], item["Điểm Danh Lần Đầu"], item["Rời Khỏi Lần Cuối"], item["Tổng Giờ"], item["Mã Tiệc"]]);
            salaryData.font = { size: 14, name: 'Times New Roman' };
            salaryData.alignment = { horizontal: "center" };
        });

        worksheet.addRow([]);

        // Get the range of cells
        const nextRow = salaryDataExport.length;
        let start = { row: 8, col: 1 },
            end = { row: nextRow, col: 5 };
        // console.log("nextRow: ", nextRow);

        const totalHour = worksheet.addRow(["Tổng Cộng", "", "", salaryDataInputedExport.totalTime]);
        totalHour.font = { bold: true, size: 14, name: 'Times New Roman' };
        totalHour.alignment = { horizontal: "center" };
        worksheet.mergeCells(`A${end.row + start.row + 2}:C${end.row + start.row + 2}`);

        const unitPrice = worksheet.addRow(["Đơn Giá", "", "", salaryDataInputedExport.unitPrice]);
        unitPrice.font = { bold: true, size: 14, name: 'Times New Roman' };
        unitPrice.alignment = { horizontal: "center" };
        worksheet.mergeCells(`A${end.row + start.row + 3}:C${end.row + start.row + 3}`);
        const unitPriceFormat = worksheet.getCell(end.row + start.row + 3, 4);
        unitPriceFormat.numFmt = '#,##0 ₫';

        const totalSalary = worksheet.addRow(["Thành Tiền", "", "", salaryDataInputedExport.totalSalary]);
        totalSalary.font = { bold: true, size: 14, name: 'Times New Roman' };
        totalSalary.alignment = { horizontal: "center" };
        worksheet.mergeCells(`A${end.row + start.row + 4}:C${end.row + start.row + 4}`);
        const totalSalaryFormat = worksheet.getCell(end.row + start.row + 4, 4);
        totalSalaryFormat.numFmt = '#,##0 ₫';
        worksheet.addRow([]);

        const dateCreate = worksheet.addRow(["", "", "", "", `Cần Thơ, ngày ${day} tháng ${month} năm ${year}`]);
        dateCreate.font = { bold: true, size: 14, name: 'Times New Roman', italic: true };
        dateCreate.alignment = { horizontal: "center" };

        const signature = worksheet.addRow(["", "", "", "", `LẬP BIỂU`]);
        signature.font = { bold: true, size: 14, name: 'Times New Roman' };
        signature.alignment = { horizontal: "center" };


        for (let i = start.row; i <= end.row + start.row + 4; i++) {
            for (let j = start.col; j <= end.col; j++) {
                worksheet.getCell(i, j).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            }
        }
    }

    // Thiết lập header và type cho response
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=' + `${fileName}.xlsx`);

    // Xuất Workbook ra response
    workbook.xlsx.write(res)
        .then(function () {
            if (type == 'detail') {
                updateCheckExported(salaryDataExport);
            }
            res.end();
        })
        .catch(function (error) {
            return res.status(500).send(error);
        });
}

let updateCheckExported = async (salaryDataExport) => {
    try {
        // Đặt múi giờ cho server
        moment.tz.setDefault("Asia/Ho_Chi_Minh");

        // Lấy thời gian hiện tại theo múi giờ đã đặt
        let currentDate = moment().tz("Asia/Ho_Chi_Minh");

        const format = "DD/MM/YYYY HH:mm:ss";
        let formatedDate = moment(currentDate).format(format);

        salaryDataExport.forEach(async (record) => {
            if (record["Đã Xuất Báo Cáo"] === undefined) {
                record["Số Điện Thoại"] = `'${record["Số Điện Thoại"]}`;
                record["Đã Xuất Báo Cáo"] = "X";
                record["Thời Gian Xuất Báo Cáo"] = formatedDate;
                await record.save();
            }
            // else {
            //     console.log("Trường này đã được cập nhật!!!");
            // }
        });
    } catch (e) {
        // console.log("lỗi cập nhật!!!")
        // const errorMessage = `Điểm danh (check-out) thất bại!`;
        // res.render("loginpage.ejs", { loginStatus: errorMessage, code: req.query.code });
    }
};

let groupBy = function (xs, key) {
    return xs.reduce(function (rv, x) {
        (rv[x[key]] = rv[x[key]] || []).push(x);
        return rv;
    }, {});
};

let currencyFormat = (money) => {
    let formattedPrice = new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' }).format(money);
    return formattedPrice.replace('₫', '').trim(); // Xóa ký tự không phải số và khoảng trắng
}

let convertCommaToDot = (totalHoursString) => {
    return totalHoursString.replace(",", ".");
}

let formatHourMinutes = (totalHours) => {
    let hours = Math.floor(totalHours);
    let minutes = Math.round((totalHours - hours) * 60);
    return hours + "h" + minutes + "'";
}

const morningStartTime = 8;
const morningEndTime = 13;
const afternoonStartTime = 13;
const afternoonEndTime = 18;
const eveningStartTime = 18;
const eveningEndTime = 23;

let getTimeSlot = (currentDate) => {
    // Tính toán thời gian hiện tại là ca nào
    let timeSlot = '';
    let hour = moment(currentDate).hour();
    if (hour >= morningStartTime && hour < morningEndTime) {
        timeSlot = 'Ca sáng';
    } else if (hour >= afternoonStartTime && hour < afternoonEndTime) {
        timeSlot = 'Ca chiều';
    } else if (hour >= eveningStartTime && hour < eveningEndTime) {
        timeSlot = 'Ca tối';
    } else {
        if (hour <= morningStartTime && hour < morningEndTime) {
            timeSlot = 'Ca sáng';
        } else if (hour <= afternoonStartTime && hour < afternoonEndTime) {
            timeSlot = 'Ca chiều';
        } else if (hour <= eveningStartTime && hour < eveningEndTime) {
            timeSlot = 'Ca tối';
        }
    }

    return timeSlot;
}

let getTimeSlotReal = (currentDate) => {
    // Tính toán thời gian hiện tại là ca nào
    let timeSlot = '';
    let hour = moment(currentDate).hour();
    if (hour >= morningStartTime && hour < morningEndTime) {
        timeSlot = 'Ca sáng';
    } else if (hour >= afternoonStartTime && hour < afternoonEndTime) {
        timeSlot = 'Ca chiều';
    } else if (hour >= eveningStartTime && hour < eveningEndTime) {
        timeSlot = 'Ca tối';
    }

    return timeSlot;
}


let roundTimeWorking = (currentDate) => {
    // Tính toán thời gian hiện tại là ca nào
    let timeStart = '';
    let hour = moment(currentDate).get('hour');
    let date = moment(currentDate).format('YYYY-MM-DD');

    const format = "DD/MM/YYYY HH:mm:ss";
    let formatedDate = moment(currentDate).format(format);

    // Làm tròn khung giờ
    if (hour >= morningStartTime && hour < morningEndTime) {
        if (hour < morningStartTime) {
            timeStart = moment(`${date} 0${morningStartTime}:00:00`).format('YYYY-MM-DD HH:mm:ss');
        } else {
            timeStart = formatedDate;
        }
    } else if (hour >= afternoonStartTime && hour < afternoonEndTime) {
        if (hour < afternoonStartTime) {
            timeStart = moment(`${date} 0${afternoonStartTime}:00:00`).format('YYYY-MM-DD HH:mm:ss');
        } else {
            timeStart = formatedDate;
        }
    } else if (hour >= eveningStartTime && hour < eveningEndTime) {
        if (hour < eveningStartTime) {
            timeStart = moment(`${date} 0${eveningStartTime}:00:00`).format('YYYY-MM-DD HH:mm:ss');
        } else {
            timeStart = formatedDate;
        }
    } else {
        if (hour < morningStartTime) {
            timeStart = moment(`${date} 0${morningStartTime}:00:00`).format('YYYY-MM-DD HH:mm:ss');
        } else if (hour < afternoonStartTime) {
            timeStart = moment(`${date} 0${afternoonStartTime}:00:00`).format('YYYY-MM-DD HH:mm:ss');
        } else if (hour < eveningStartTime) {
            timeStart = moment(`${date} 0${eveningStartTime}:00:00`).format('YYYY-MM-DD HH:mm:ss');
        } else {
            console.log("ko the diem danh luc nay")
        }
    }

    return timeStart;
}


module.exports = {
    getListUsers: getListUsers,
    getLoginpage: getLoginpage,
    login: login,
    getTimekeepingPage: getTimekeepingPage,
    getQRpage: getQRpage,
    createQR: createQR,
    getSalaryPage: getSalaryPage,
    salary: salary,
    exportExcel: exportExcel
};
