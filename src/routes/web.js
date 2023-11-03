import express from "express";
import homepageController from "../controllers/homepageController";

let router = express.Router();

let initWebRoutes = (app) => {
    // router.get("/", homepageController.getHomepage);
    router.get("/chamcong/getlistusers", homepageController.getListUsers);
    router.get("/chamcong", homepageController.getLoginpage);
    router.post("/chamcong", homepageController.login);
    router.get("/chamcong/qr", homepageController.getQRpage);
    router.post("/chamcong/qr", homepageController.createQR);

    // router.get("/excel", homepageController.insertGoogleSheet);
    // router.get("/update", homepageController.updateGoogleSheet);
    
    router.get("/chamcong", homepageController.getTimekeepingPage);
    router.get("/chamcong/report/:type", homepageController.getSalaryPage);
    router.get("/chamcong/report", homepageController.getSalaryPage);

    router.post("/chamcong/report/:type", homepageController.salary);
    router.get("/chamcong/export-excel/:type", homepageController.exportExcel);

    return app.use("/", router);
};

module.exports = initWebRoutes;