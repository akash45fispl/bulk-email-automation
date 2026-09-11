const xlsx = require('xlsx');
const path = require('path');

const sampleClients = [
  { Name: 'John Smith', Email: 'john.smith@example.com', Company: 'Apex Innovations', City: 'New York', AccountID: 'ACC-1001' },
  { Name: 'Sarah Connor', Email: 'sarah.c@cyberdyne.io', Company: 'Cyberdyne Systems', City: 'San Francisco', AccountID: 'ACC-1002' },
  { Name: 'Michael Scott', Email: 'michael@dundermifflin.com', Company: 'Dunder Mifflin', City: 'Scranton', AccountID: 'ACC-1003' },
  { Name: 'Elena Rostova', Email: 'elena.rostova@techglobal.io', Company: 'TechGlobal Inc', City: 'London', AccountID: 'ACC-1004' },
  { Name: 'David Chen', Email: 'dchen@innovatecorp.com', Company: 'Innovate Corp', City: 'Toronto', AccountID: 'ACC-1005' },
  { Name: 'Priya Sharma', Email: 'priya.s@cloudscale.net', Company: 'CloudScale Solutions', City: 'Bengaluru', AccountID: 'ACC-1006' },
  { Name: 'Carlos Mendez', Email: 'cmendez@latamventures.co', Company: 'LatAm Ventures', City: 'Madrid', AccountID: 'ACC-1007' },
  { Name: 'Aisha Al-Mansoor', Email: 'aisha@emiratesfuture.ae', Company: 'Emirates Future Hub', City: 'Dubai', AccountID: 'ACC-1008' }
];

const ws = xlsx.utils.json_to_sheet(sampleClients);
const wb = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(wb, ws, 'Clients');

const outPath = path.join(__dirname, 'sample_clients.xlsx');
xlsx.writeFile(wb, outPath);
console.log('Created sample_clients.xlsx successfully at:', outPath);
