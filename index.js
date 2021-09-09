const express = require("express");
const { exec } = require("child_process");
const app = new express();
const _PORT = 4022;

app.use(express.json());

app.listen(_PORT, () => {
    console.log("Listening on port 4022");
});

app.get('/closeExcel', (req, res) => {
    exec("taskkill /IM excel.exe");
    res.send("Excel closed on 172.16.213.254");
});