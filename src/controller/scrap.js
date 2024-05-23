const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

// Kabupaten Kota Indonesia
const getKabupatenKota = async (req, res) => {
  const typeFile = req.body.typeFile || "xlsx";
  const fileName = "kabupaten-kota-indonesia";
  const fileDir = path.join(__dirname, "../../", "public", "export", `${fileName}.${typeFile}`);

  res.setHeader("Content-Disposition", "attachment; filename=" + `${fileName}.${typeFile}`);
  res.setHeader("Content-Transfer-Encoding", "binary");
  res.setHeader("Content-Type", "application/octet-stream");

  if (!fs.existsSync(fileDir)) {
    try {
      // Membuka browser
      const browser = await puppeteer.launch({
        headless: true,
      });

      const page = await browser.newPage();

      // Arahkan ke URL
      await page.goto("https://id.wikipedia.org/wiki/Daftar_kabupaten_dan_kota_di_Indonesia");

      // Kolom Header
      let headers = [];
      let numberTable = 16;

      while (headers.length === 0 && numberTable < 200) {
        headers = await page.$$eval(`#mw-content-text > div.mw-content-ltr.mw-parser-output > table:nth-child(${numberTable}) > thead > tr th`, (elements) => {
          return elements.map((el) => el.innerText.trim().toLowerCase().replace(".", ""));
        });

        numberTable++;
      }

      await headers.pop();
      await headers.pop();

      // Inisialisasi Excel JS
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet("Kabupatan-Kota-Indonesia");

      // Menambahkan Kolom Header
      await sheet.addRow(headers);

      // Menambahkan Data
      let number = 1;
      for (let i = 3; i < 205; i++) {
        const data = await page.$$eval(`#mw-content-text > div.mw-content-ltr.mw-parser-output > table:nth-child(${i}) > tbody > tr`, (elements) => {
          return elements.map((el) => el.innerText.trim("\n").split("\t"));
        });

        for (let j = 0; j < data.length; j++) {
          data[j][0] = number;
          await sheet.addRow(data[j]);
          number++;
        }
      }

      if (typeFile == "csv") {
        await workbook.csv.writeFile(fileDir);
      } else {
        await workbook.xlsx.writeFile(fileDir);
      }

      await browser.close();

      // Return File
      res.sendFile(fileDir, function (err) {
        if (err) {
          res.status(404).json({
            path: fileDir,
            err,
          });
        }
      });
    } catch (error) {
      res.status(500).json({ error });
    }
  } else {
    res.sendFile(fileDir, function (err) {
      if (err) {
        res.status(404).json({
          path: fileDir,
          err,
        });
      }
    });
  }
};

// Kabupaten Kota By Provinsi
const getKabupatenKotaByProvinsi = async (req, res) => {
  const typeFile = req.body.typeFile || "xlsx";
  let fileName = req.body.provinsi;
  const fileDir = path.join(__dirname, "../../", "public", "export", `${fileName}.${typeFile}`);

  res.setHeader("Content-Disposition", "attachment; filename=" + `${fileName}.${typeFile}`);
  res.setHeader("Content-Transfer-Encoding", "binary");
  res.setHeader("Content-Type", "application/octet-stream");

  if (!fs.existsSync(fileDir)) {
    try {
      // Membuka browser
      const browser = await puppeteer.launch({
        headless: true,
      });

      const page = await browser.newPage();

      // Arahkan ke URL
      await page.goto("https://id.wikipedia.org/wiki/Daftar_kabupaten_dan_kota_di_Indonesia");

      // Mencari Provinsi dan menyimpan ke dalam json file
      const jsonDir = path.join(__dirname, "../../", "public", "provinsi.json");

      let provinsi = [];

      // Menulis File JSON
      if (!fs.existsSync(jsonDir)) {
        for (let i = 7; i < 200; i++) {
          const data = await page.$$eval(`#mw-content-text > div.mw-content-ltr.mw-parser-output > h3:nth-child(${i}) span`, (elements) => {
            return elements.map((el) => el.innerText.trim());
          });
          if (data[0] !== undefined) {
            provinsi.push([data[0].replace(/ /g, "_").toLowerCase(), i]);
          }
        }

        await fs.writeFile(jsonDir, JSON.stringify(provinsi, null, 2), (err, data) => {
          if (err) {
            console.log("Error", err);
            return;
          }
        });
      }

      // Membaca File JSON
      const dataJson = await fs.promises.readFile(jsonDir, "utf8");
      provinsi = JSON.parse(dataJson);

      // Mencari Provinsi Berdasarkan Parameter
      const find = provinsi.find((item) => item[0] === fileName);

      // Inisialisasi ExcelJS
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet(`${find[0]}`);

      // Kolom Header
      const headers = await page.$$eval(`#mw-content-text > div.mw-content-ltr.mw-parser-output > table:nth-child(${find[1] + 4}) > thead > tr th`, (elements) => {
        return elements.map((el) => el.innerText.trim().toLowerCase().replace(".", ""));
      });

      await headers.pop();
      await headers.pop();

      await sheet.addRow(headers);

      // Data
      const data = await page.$$eval(`#mw-content-text > div.mw-content-ltr.mw-parser-output > table:nth-child(${find[1] + 4}) > tbody > tr`, (elements) => {
        return elements.map((el) => el.innerText.trim("\n").split("\t"));
      });

      for (let j = 0; j < data.length; j++) {
        await sheet.addRow(data[j]);
      }

      if (typeFile == "csv") {
        await workbook.csv.writeFile(fileDir);
      } else {
        await workbook.xlsx.writeFile(fileDir);
      }

      await browser.close();

      // Return File
      res.sendFile(fileDir, function (err) {
        if (err) {
          res.status(404).json({
            path: fileDir,
            err,
          });
        }
      });
    } catch (error) {
      throw error;
    }
  } else {
    res.sendFile(fileDir, function (err) {
      if (err) {
        res.status(404).json({
          path: fileDir,
          err,
        });
      }
    });
  }
};

// Provinsi Indonesia
const getProvinsi = async (req, res) => {
  const typeFile = req.body.typeFile || "xlsx";
  const fileName = "provinsi";
  const fileDir = path.join(__dirname, "../../", "public", "export", `${fileName}.${typeFile}`);

  res.setHeader("Content-Disposition", "attachment; filename=" + `${fileName}.${typeFile}`);
  res.setHeader("Content-Transfer-Encoding", "binary");
  res.setHeader("Content-Type", "application/octet-stream");

  if (!fs.existsSync(fileDir)) {
    try {
      // Membuka browser
      const browser = await puppeteer.launch({
        headless: true,
      });

      const page = await browser.newPage();

      // Arahkan ke URL
      await page.goto("https://id.wikipedia.org/wiki/Daftar_kabupaten_dan_kota_di_Indonesia");

      // Kolom Header
      const headers = await page.$$eval(`#mw-content-text > div.mw-content-ltr.mw-parser-output > table:nth-child(2) > thead > tr th`, (elements) => {
        return elements.map((el) => el.innerText.trim().toLowerCase().replace(".", ""));
      });

      await headers.splice(1, 1);

      // Menambahkan data
      const data = await page.$$eval(`#mw-content-text > div.mw-content-ltr.mw-parser-output > table:nth-child(2) > tbody > tr`, (elements) => {
        return elements.map((el) => el.innerText.trim("\n").split("\t"));
      });

      // Inisialisasi Excel JS
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet("Kabupatan-Kota-Indonesia");

      // Menambahkan Kolom Header
      await sheet.addRow(headers);

      for (let j = 0; j < data.length - 1; j++) {
        if (data[j].length === 6) {
          data[j].splice(1, 1);
        }
        await sheet.addRow(data[j]);
      }

      if (typeFile == "csv") {
        await workbook.csv.writeFile(fileDir);
      } else {
        await workbook.xlsx.writeFile(fileDir);
      }

      await browser.close();

      // Return File
      res.sendFile(fileDir, function (err) {
        if (err) {
          res.status(404).json({
            path: fileDir,
            err,
          });
        }
      });
    } catch (error) {
      throw error;
    }
  } else {
    res.sendFile(fileDir, function (err) {
      if (err) {
        res.status(404).json({
          path: fileDir,
          err,
        });
      }
    });
  }
};

const provinsiJson = async (req, res) => {
  try {
    const jsonDir = path.join(__dirname, "../../", "public", "provinsi.json");

    let provinsi = [];

    if (!fs.existsSync(jsonDir)) {
      // Membuka browser
      const browser = await puppeteer.launch({
        headless: true,
      });

      const page = await browser.newPage();

      // Arahkan ke URL
      await page.goto("https://id.wikipedia.org/wiki/Daftar_kabupaten_dan_kota_di_Indonesia");

      // Mencari Provinsi dan menyimpan ke dalam json file
      // Menulis File JSON
      for (let i = 7; i < 200; i++) {
        const data = await page.$$eval(`#mw-content-text > div.mw-content-ltr.mw-parser-output > h3:nth-child(${i}) span`, (elements) => {
          return elements.map((el) => el.innerText.trim());
        });
        if (data[0] !== undefined) {
          provinsi.push([data[0].replace(/ /g, "_").toLowerCase(), i]);
        }
      }

      await fs.writeFile(jsonDir, JSON.stringify(provinsi, null, 2), (err, data) => {
        if (err) {
          console.log("Error", err);
          return;
        }
      });

      await browser.close();
    }

    // Membaca File JSON
    const dataJson = await fs.promises.readFile(jsonDir, "utf8");
    provinsi = JSON.parse(dataJson);

    res.render("index", {
      layout: "layout/main",
      title: "Scrapper",
      provinsi,
    });
  } catch (error) {
    res.send(500).json({ error });
  }
};
module.exports = { getKabupatenKota, getProvinsi, getKabupatenKotaByProvinsi, provinsiJson };
