const express = require('express');
const xlsx = require('xlsx');
const app = express();
const port = 3000;

app.use(express.static('public')); // Serve static files from 'public' directory

// 提供 DRAM 模型的下拉選單列表
app.get('/models', (req, res) => {
    try {
        const workbook = xlsx.readFile('Jetway AVL.xlsx');
        const worksheet = workbook.Sheets['DRAM'];
        if (!worksheet) {
            return res.status(404).send('DRAM models not found');
        }
        const data = xlsx.utils.sheet_to_json(worksheet);
        const models = data.map(row => row.Model);
        res.json(models);
    } catch (error) {
        console.error('Error reading the Excel file:', error);
        res.status(500).send('Error reading the Excel file');
    }
});

// 提供 SATA 模型的列表下拉選單列表
app.get('/sata-models', (req, res) => {
    try {
        const workbook = xlsx.readFile('Jetway AVL.xlsx');
        const worksheet = workbook.Sheets['SATA'];
        if (!worksheet) {
            return res.status(404).send('SATA models not found');
        }
        const data = xlsx.utils.sheet_to_json(worksheet);
        const models = data.map(row => row.Model);
        res.json(models);
    } catch (error) {
        console.error('Error reading the Excel file:', error);
        res.status(500).send('Error reading the Excel file');
    }
});

// 提供 mSATA 模型的列表下拉菜单列表
app.get('/msata-models', (req, res) => {
    try {
        const workbook = xlsx.readFile('Jetway AVL.xlsx');
        const worksheet = workbook.Sheets['mSATA']; // 确认这里的 'mSATA' 与 Excel 中的工作表名称匹配
        if (!worksheet) {
            return res.status(404).send('mSATA models not found');
        }
        const data = xlsx.utils.sheet_to_json(worksheet);
        const models = data.map(row => row.Model);
        res.json(models);
    } catch (error) {
        console.error('Error reading the Excel file:', error);
        res.status(500).send('Error reading the Excel file');
    }
});

// 提供 m2SATA 模型的列表下拉菜单列表
app.get('/m2sata-models', (req, res) => {
    try {
        const workbook = xlsx.readFile('Jetway AVL.xlsx');
        const worksheet = workbook.Sheets['M2SATA']; // 确认这里的 'm2SATA' 与 Excel 中的工作表名称匹配
        if (!worksheet) {
            return res.status(404).send('m2SATA models not found');
        }
        const data = xlsx.utils.sheet_to_json(worksheet);
        const models = data.map(row => row.Model);
        res.json(models);
    } catch (error) {
        console.error('Error reading the Excel file:', error);
        res.status(500).send('Error reading the Excel file');
    }
});

// 提供 m2PCIE 模型的列表下拉菜单列表
app.get('/m2pcie-models', (req, res) => {
    try {
        const workbook = xlsx.readFile('Jetway AVL.xlsx');
        const worksheet = workbook.Sheets['M2PCIE']; // 确认这里的 'M2PCIE' 与 Excel 中的工作表名称匹配
        if (!worksheet) {
            return res.status(404).send('m2pcie models not found');
        }
        const data = xlsx.utils.sheet_to_json(worksheet);
        const models = data.map(row => row.Model);
        res.json(models);
    } catch (error) {
        console.error('Error reading the Excel file:', error);
        res.status(500).send('Error reading the Excel file');
    }
});

// 提供 CPU 模型的列表下拉菜单列表
app.get('/cpu-models', (req, res) => {
    try {
        const workbook = xlsx.readFile('Jetway AVL.xlsx');
        const worksheet = workbook.Sheets['CPU']; // 确认这里的 'CPU' 与 Excel 中的工作表名称匹配
        if (!worksheet) {
            return res.status(404).send('CPU models not found');
        }
        const data = xlsx.utils.sheet_to_json(worksheet);
        const models = data.map(row => row.Model);
        res.json(models);
    } catch (error) {
        console.error('Error reading the Excel file:', error);
        res.status(500).send('Error reading the Excel file');
    }
});

// 根据模型获取 DRAM 详细信息的端点
app.get('/model-details', (req, res) => {
    const model = req.query.model;
    try {
        const workbook = xlsx.readFile('Jetway AVL.xlsx');
        const worksheet = workbook.Sheets['DRAM'];
        if (!worksheet) {
            return res.status(404).send('DRAM sheet not found');
        }
        const data = xlsx.utils.sheet_to_json(worksheet);
        const modelDetails = data.find(row => row.Model === model);
        
        if (modelDetails) {
            res.json(modelDetails);
        } else {
            res.status(404).send('Model not found in DRAM');
        }
    } catch (error) {
        console.error('Error processing the Excel file:', error);
        res.status(500).send('Error processing the Excel file');
    }
});

// 根据模型获取 SATA 详细信息的端点
app.get('/sata-details', (req, res) => {
    const model = req.query.model;
    try {
        const workbook = xlsx.readFile('Jetway AVL.xlsx');
        const worksheet = workbook.Sheets['SATA'];
        if (!worksheet) {
            return res.status(404).send('SATA sheet not found');
        }
        const data = xlsx.utils.sheet_to_json(worksheet);
        const modelDetails = data.find(row => row.Model === model);
        
        if (modelDetails) {
            res.json(modelDetails);
        } else {
            res.status(404).send('Model not found in SATA');
        }
    } catch (error) {
        console.error('Error processing the Excel file:', error);
        res.status(500).send('Error processing the Excel file');
    }
});
// 根据模型获取 mSATA 详细信息的端点
app.get('/msata-details', (req, res) => {
    const model = req.query.model;
    try {
        const workbook = xlsx.readFile('Jetway AVL.xlsx');
        const worksheet = workbook.Sheets['mSATA']; // 确认这里的 'mSATA' 与 Excel 中的工作表名称匹配
        if (!worksheet) {
            return res.status(404).send('mSATA sheet not found');
        }
        const data = xlsx.utils.sheet_to_json(worksheet);
        const modelDetails = data.find(row => row.Model === model);
        
        if (modelDetails) {
            res.json(modelDetails);
        } else {
            res.status(404).send('Model not found in mSATA');
        }
    } catch (error) {
        console.error('Error processing the Excel file:', error);
        res.status(500).send('Error processing the Excel file');
    }
});
// 根据模型获取 m2SATA 详细信息的端点
app.get('/m2sata-details', (req, res) => {
    const model = req.query.model;
    try {
        const workbook = xlsx.readFile('Jetway AVL.xlsx');
        const worksheet = workbook.Sheets['M2SATA']; // 确认这里的 'M2SATA' 与 Excel 中的工作表名称匹配
        if (!worksheet) {
            return res.status(404).send('M2SATA sheet not found');
        }
        const data = xlsx.utils.sheet_to_json(worksheet);
        const modelDetails = data.find(row => row.Model === model);
        
        if (modelDetails) {
            res.json(modelDetails);
        } else {
            res.status(404).send('Model not found in M2SATA');
        }
    } catch (error) {
        console.error('Error processing the Excel file:', error);
        res.status(500).send('Error processing the Excel file');
    }
});

// 根据模型获取 m2PCIE 详细信息的端点
app.get('/m2pcie-details', (req, res) => {
    const model = req.query.model;
    try {
        const workbook = xlsx.readFile('Jetway AVL.xlsx');
        const worksheet = workbook.Sheets['M2PCIE']; // 确认这里的 'M2PCIE' 与 Excel 中的工作表名称匹配
        if (!worksheet) {
            return res.status(404).send('M2PCIE sheet not found');
        }
        const data = xlsx.utils.sheet_to_json(worksheet);
        const modelDetails = data.find(row => row.Model === model);
        
        if (modelDetails) {
            res.json(modelDetails);
        } else {
            res.status(404).send('Model not found in M2PCIE');
        }
    } catch (error) {
        console.error('Error processing the Excel file:', error);
        res.status(500).send('Error processing the Excel file');
    }
});

// 根据模型获取 CPU 详细信息的端点
app.get('/cpu-details', (req, res) => {
    const model = req.query.model;
    try {
        const workbook = xlsx.readFile('Jetway AVL.xlsx');
        const worksheet = workbook.Sheets['CPU']; // 确认这里的 'CPU' 与 Excel 中的工作表名称匹配
        if (!worksheet) {
            return res.status(404).send('CPU sheet not found');
        }
        const data = xlsx.utils.sheet_to_json(worksheet);
        const modelDetails = data.find(row => row.Model === model);
        
        if (modelDetails) {
            res.json(modelDetails);
        } else {
            res.status(404).send('Model not found in CPU');
        }
    } catch (error) {
        console.error('Error processing the Excel file:', error);
        res.status(500).send('Error processing the Excel file');
    }
});
app.listen(3000, '0.0.0.0', () => {
    console.log('Server is running on http://localhost:3000');
});
