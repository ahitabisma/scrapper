const express = require("express");
const expressLayouts = require("express-ejs-layouts");
const path = require("path");
const { getKabupatenKota, getProvinsi, getKabupatenKotaByProvinsi, provinsiJson } = require("./controller/scrap");
const favicon = require("serve-favicon");
const app = express();
const port = 3000;

app.use(favicon(path.join(__dirname, "../", "public", "img", "logo.png")));
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(expressLayouts);
app.use(express.static(path.join(__dirname, "public")));
app.use(express.urlencoded({ extended: true }));

// Halaman Index
app.get("/", provinsiJson);

// Kabupaten Kota
app.post("/kabupaten-kota-indonesia", getKabupatenKota);

// Kabupaten Kota By Provinsi
app.post("/kabupaten-kota", getKabupatenKotaByProvinsi);

// Provinsi
app.post("/provinsi", getProvinsi);

app.listen(port, () => {
  console.log(`Scrap app listening on port ${port}`);
});
