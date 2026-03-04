const express = require("express");
const cors = require("cors");
const path = require("path");
const reportRoutes = require("./routes/reportRoutes");

const app = express();

app.use(cors());
app.use(express.json());
app.use("/api/report", reportRoutes);

// serve frontend
app.use(express.static(path.join(__dirname, "../frontend")));

app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "../frontend/index.html"));
});

app.listen(5000, () => {
  console.log("Server running at http://localhost:5000");
});