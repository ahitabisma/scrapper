const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

const getKabupatenKota = async (req, res) => {
  // Membuka browser
  const browser = await puppeteer.launch({
    headless: true,
  });

  try {
    const page = await browser.newPage();

    // Arahkan ke URL
    await page.goto("https://id.wikipedia.org/wiki/Daftar_kabupaten_dan_kota_di_Indonesia");

    // Kolom Header
    const headers = await page.$$eval(`#mw-content-text > div.mw-content-ltr.mw-parser-output > table:nth-child(18) > thead > tr th`, (elements) => {
      return elements.map((el) => el.innerText.trim().toLowerCase().replace(".", ""));
    });

    await headers.pop();
    await headers.pop();
    // console.log(headers);

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

    const filePath = path.join(__dirname, "../../", "public", "export", "kabupaten-kota-indonesia.xlsx");
    const filePathCsv = path.join(__dirname, "../../", "public", "export", "kabupaten-kota-indonesia.csv");

    // Cek File Excel
    if (!fs.existsSync(filePath)) {
      await workbook.xlsx.writeFile(filePath);
      console.log(`Excel file saved to: ${filePath}`);
      res.send({ path: filePath, filePathCsv: filePathCsv });
    } else {
      console.log("File Excel sudah ada");
      res.send({ path: filePath, filePathCsv: filePathCsv });
      // res.sendFile(filePath);
    }

    // Cek File CSV
    if (!fs.existsSync(filePathCsv)) {
      await workbook.csv.writeFile(filePathCsv);
      console.log(`Excel file saved to: ${filePathCsv}`);
      res.send({ path: filePath, filePathCsv: filePathCsv });
      // res.sendFile(filePathCsv);
    } else {
      console.log("File CSV sudah ada");
      res.send({ path: filePath, filePathCsv: filePathCsv });
      // res.sendFile(filePathCsv);
    }

    await browser.close();
  } catch (error) {
    console.error(error);
    await browser.close();
  }
};

const getKabKotaByProvinsi = async (prov) => {
  // Membuka browser
  const browser = await puppeteer.launch({
    headless: true,
  });

  try {
    const page = await browser.newPage();

    // Arahkan ke URL
    await page.goto("https://id.wikipedia.org/wiki/Daftar_kabupaten_dan_kota_di_Indonesia");

    let provinsi = [];
    for (let i = 7; i < 200; i++) {
      const data = await page.$$eval(`#mw-content-text > div.mw-content-ltr.mw-parser-output > h3:nth-child(${i}) span`, (elements) => {
        return elements.map((el) => el.innerText.trim());
      });
      if (data[0] !== undefined) {
        provinsi.push([data[0].replace(/ /g, "_").toLowerCase(), i]);
      }
    }

    const val = prov.toLowerCase().replace(/ /g, "_");

    //   Mencari Provinsi Berdasarkan Parameter
    const find = provinsi.find((item) => item[0] === val);

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(`${find[0]}`);

    const headers = await page.$$eval(`#mw-content-text > div.mw-content-ltr.mw-parser-output > table:nth-child(${find[1] + 4}) > thead > tr th`, (elements) => {
      return elements.map((el) => el.innerText.trim().toLowerCase().replace(".", ""));
    });

    await headers.splice(headers.length, 1);
    await headers.splice(headers.length - 1, 1);

    await sheet.addRow(headers);

    const data = await page.$$eval(`#mw-content-text > div.mw-content-ltr.mw-parser-output > table:nth-child(${find[1] + 4}) > tbody > tr`, (elements) => {
      return elements.map((el) => el.innerText.trim("\n").split("\t"));
    });

    for (let j = 0; j < data.length; j++) {
      await sheet.addRow(data[j]);
    }

    // const filePath = `./src/export/${find[0]}.xlsx`;
    // const filePathCsv = `./src/export/${find[0]}.csv`;
    const filePath = path.join(__dirname, "../../", "public", "export", `${find[0]}.xlsx`);
    const filePathCsv = path.join(__dirname, "../../", "public", "export", `${find[0]}.csv`);

    // Cek File Excel
    if (!fs.existsSync(filePath)) {
      await workbook.xlsx.writeFile(filePath);
      console.log(`Excel file saved to: ${filePath}`);
    } else {
      console.log("File Excel sudah ada");
    }

    // Cek File CSV
    if (!fs.existsSync(filePathCsv)) {
      await workbook.csv.writeFile(filePathCsv);
      console.log(`Excel file saved to: ${filePathCsv}`);
    } else {
      console.log("File CSV sudah ada");
    }

    await browser.close();
  } catch (error) {
    console.log("Provinsi tidak ditemukan");
    await browser.close();
  }
};

const getProvinsi = async () => {
  // Membuka browser
  const browser = await puppeteer.launch({
    headless: true,
  });

  try {
    const page = await browser.newPage();

    // Arahkan ke URL
    await page.goto("https://id.wikipedia.org/wiki/Daftar_kabupaten_dan_kota_di_Indonesia");

    //   Kolom Header
    const headers = await page.$$eval(`#mw-content-text > div.mw-content-ltr.mw-parser-output > table:nth-child(2) > thead > tr th`, (elements) => {
      return elements.map((el) => el.innerText.trim().toLowerCase().replace(".", ""));
    });

    await headers.splice(1, 1);

    const data = await page.$$eval(`#mw-content-text > div.mw-content-ltr.mw-parser-output > table:nth-child(2) > tbody > tr`, (elements) => {
      return elements.map((el) => el.innerText.trim("\n").split("\t"));
    });

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Provinsi");

    await sheet.addRow(headers);

    for (let j = 0; j < data.length - 1; j++) {
      if (data[j].length === 6) {
        data[j].splice(1, 1);
      }
      await sheet.addRow(data[j]);
    }

    // const filePath = `./src/export/provinsi.xlsx`;
    // const filePathCsv = `./src/export/provinsi.csv`;
    const filePath = path.join(__dirname, "../", "public", "export", "provinsi.xlsx");
    const filePathCsv = path.join(__dirname, "../", "public", "export", "provinsi.csv");

    // Cek File Excel
    if (!fs.existsSync(filePath)) {
      await workbook.xlsx.writeFile(filePath);
      console.log(`Excel file saved to: ${filePath}`);
    } else {
      console.log("File Excel sudah ada");
    }

    // Cek File CSV
    if (!fs.existsSync(filePathCsv)) {
      await workbook.csv.writeFile(filePathCsv);
      console.log(`Excel file saved to: ${filePathCsv}`);
    } else {
      console.log("File CSV sudah ada");
    }

    await browser.close();
  } catch (error) {
    console.error(error);
    await browser.close();
  }
};

module.exports = { getKabupatenKota, getKabKotaByProvinsi, getProvinsi };
