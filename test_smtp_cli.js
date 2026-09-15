const nodemailer = require('nodemailer');
const net = require('net');

async function testNetwork(host, port) {
  return new Promise((resolve) => {
    console.log(`[1] Testing network connection to ${host}:${port}...`);
    const socket = net.createConnection(port, host, () => {
      console.log(`✅ Network reachable: Successfully connected to ${host}:${port}`);
      socket.destroy();
      resolve(true);
    });
    socket.setTimeout(5000);
    socket.on('timeout', () => {
      console.log(`❌ Network timeout: Could not reach ${host}:${port}`);
      socket.destroy();
      resolve(false);
    });
    socket.on('error', (err) => {
      console.log(`❌ Network error on ${host}:${port}: ${err.message}`);
      resolve(false);
    });
  });
}

async function testAuth(user, pass) {
  console.log(`\n[2] Testing Gmail SMTP Authentication for user: ${user}...`);
  const cleanPass = String(pass || '').replace(/\s+/g, '');
  
  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
      user: user.trim(),
      pass: cleanPass
    }
  });

  try {
    const verified = await transporter.verify();
    console.log(`✅ AUTHENTICATION SUCCESSFUL! Connection verified:`, verified);
    return true;
  } catch (err) {
    console.log(`❌ AUTHENTICATION FAILED: ${err.message}`);
    if (err.responseCode === 535 || err.message.includes('BadCredentials')) {
      console.log(`\n👉 Reason: Google rejected the password. Please verify that:`);
      console.log(`   1. 2-Step Verification is ON in your Google Account.`);
      console.log(`   2. You are using a 16-letter App Password generated from https://myaccount.google.com/apppasswords`);
    }
    return false;
  }
}

async function main() {
  const email = process.argv[2] || 'akash45.fispl@gmail.com';
  const pass = process.argv[3] || '';

  console.log('==================================================');
  console.log('       SMTP DIAGNOSTIC & CONNECTIVITY TEST        ');
  console.log('==================================================');

  await testNetwork('smtp.gmail.com', 587);
  await testNetwork('smtp.gmail.com', 465);

  if (pass) {
    await testAuth(email, pass);
  } else {
    console.log(`\nℹ️ To test authentication with your password, run:`);
    console.log(`   node test_smtp_cli.js <your_email> <your_16_letter_app_password>`);
  }
  console.log('==================================================\n');
}

main().catch(console.error);
