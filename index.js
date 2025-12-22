const express = require("express");
const axios = require("axios");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
//только локально
// require("dotenv").config();

const app = express();
const port = 3000;

app.use(express.static("public"));
app.use(express.urlencoded({ extended: true }));

const YA_API_URL = "https://cloud-api.yandex.net/v1/disk/resources";
const TOKEN = process.env.TOKEN;

const yandexApi = axios.create({
  baseURL: YA_API_URL,
  headers: {
    Authorization: `OAuth ${TOKEN}`,
    "Content-Type": "application/json",
  },
});

app.post("/add-data", (req, res) => {
  async function editExcelOnYandexDisk(filePath, newData) {
    try {
      // 1. Получить ссылку на скачивание файла
      const downloadResponse = await yandexApi.get("/download", {
        params: { path: filePath },
      });
      const downloadUrl = downloadResponse.data.href;

      // 2. Скачать файл
      const fileResponse = await axios.get(downloadUrl, {
        responseType: "arraybuffer",
      });
      const workbook = XLSX.read(fileResponse.data, { type: "buffer" });

      // 3. Изменить данные (пример: добавление в первую страницу)
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      // Преобразуйте newData в формат строк (пример)
      // newData должен быть массивом массивов [[A1, B1], [A2, B2], ...]
      XLSX.utils.sheet_add_aoa(worksheet, newData, { origin: -1 }); // -1 добавляет в конец

      // 4. Подготовить обновленный файл
      const updatedBuffer = XLSX.write(workbook, {
        type: "buffer",
        bookType: "xlsx",
      });

      // 5. Получить ссылку для загрузки
      const uploadResponse = await yandexApi.get("/upload", {
        params: {
          path: filePath,
          overwrite: true,
        },
      });
      const uploadUrl = uploadResponse.data.href;

      // 6. Загрузить обновленный файл
      await axios.put(uploadUrl, updatedBuffer, {
        headers: {
          "Content-Type":
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
      });
      console.log("Файл успешно обновлен!");
      res.sendFile(path.join(__dirname, "success.html"));
    } catch (error) {
      console.error("Ошибка:", error.response?.data || error.message);
      res.send("Ошибка:", error.message);
    }
  }

  // Пример использования
  const newData = [
    [req.body.name, req.body.thing, req.body.num, req.body.date],
  ];

  editExcelOnYandexDisk("/ЖУРНАЛ/Инвентаризация.xlsx", newData);
});

app.listen(port, () => console.log("server start"));

