const express = require("express");
const cors = require("cors");
const path = require("path");
const reportRoutes = require("./routes/reportRoutes");

const app = express();

app.use(cors());
app.use(express.json());
app.use("/api/report", reportRoutes);

// serve frontend
// app.use(express.static(path.join(__dirname, "../frontend")));

// app.get("*", (req, res) => {
//   res.sendFile(path.join(__dirname, "../frontend/index.html"));
// });

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
