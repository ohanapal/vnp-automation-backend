import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Define the data.json file path
const dataFilePath = path.join(__dirname, 'data.json');

// Initialize data.json if it doesn't exist
if (!fs.existsSync(dataFilePath)) {
    fs.writeFileSync(dataFilePath, JSON.stringify({ logs: [] }, null, 2));
}

// Create a write stream for the log file
const logStream = fs.createWriteStream(dataFilePath, { flags: 'a' });

// Custom logger function
const logger = {
    info: (message) => {
        const timestamp = new Date().toISOString();
        const logEntry = {
            timestamp,
            level: 'INFO',
            message,
            data: null
        };
        
        console.log(message);
        appendToLog(logEntry);
    },
    error: (message, error = null) => {
        const timestamp = new Date().toISOString();
        const logEntry = {
            timestamp,
            level: 'ERROR',
            message,
            data: error ? {
                name: error.name,
                message: error.message,
                stack: error.stack
            } : null
        };
        
        console.error(message);
        appendToLog(logEntry);
    },
    warn: (message) => {
        const timestamp = new Date().toISOString();
        const logEntry = {
            timestamp,
            level: 'WARN',
            message,
            data: null
        };
        
        console.warn(message);
        appendToLog(logEntry);
    },
    data: (message, data) => {
        const timestamp = new Date().toISOString();
        const logEntry = {
            timestamp,
            level: 'DATA',
            message,
            data
        };
        
        console.log(message);
        appendToLog(logEntry);
    }
};

// Helper function to append logs to data.json
function appendToLog(logEntry) {
    try {
        // Read the current content
        const fileContent = fs.readFileSync(dataFilePath, 'utf8');
        const data = JSON.parse(fileContent);
        
        // Add the new log entry
        data.logs.push(logEntry);
        
        // Write the entire updated content
        fs.writeFileSync(dataFilePath, JSON.stringify(data, null, 2));
    } catch (error) {
        console.error('Error writing to log file:', error);
    }
}

// Handle process termination
process.on('SIGINT', () => {
    logStream.end();
    process.exit();
});

export default logger; 

// 1. Instead of hotel name, we use an array containing hotel name and reservation ID array
// 2. We provide Property name in searchbar and search
// 2. Instead of Date, we provide reservation ID in search bar




// hotels = [
//     {
//         name: 'Hotel name',
//         idList: [
//             '1535123', '1655352'
//         ]
//     }
// ]

// hotels2 = [
//     {
//         name: 'yhkbkjb',
//         reservationID: '5616531'
//     },
//     {
//         name: 'yhkbkjb',
//         reservationID: '5616552'
//     }
// ]