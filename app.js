import express from 'express';
import excel from './app/excel';
// import { read } from 'fs';
const app = express();

app.get('/api', (req, res) => {
    res.status(200).send({
        success: 'true',
        message: 'CSV file written!',
        excel: excel.excel()
    })
});

app.get('/read', (req, res) => {
    res.status(200).send({
        success: 'true',
        message: 'EXCEL file read!',
        excel: excel.read()
    })
});

app.get('/write', (req, res) => {
    res.status(200).send({
        success: 'true',
        message: 'CSV file created!',
        excel: excel.write()
    })
});

const PORT = 5000;

app.listen(PORT, () => {
    console.log(`API running on port ${PORT}`)
});