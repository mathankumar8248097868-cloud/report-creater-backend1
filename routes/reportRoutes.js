const express = require("express");
const multer = require("multer");
const router = express.Router();

const upload = multer({ dest: "uploads/" });

const { generateReport } = require("../controllers/reportController");

router.post("/generate", upload.array("photos"), generateReport);

module.exports = router;