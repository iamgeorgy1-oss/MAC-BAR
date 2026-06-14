// MAC Športni Bar — Reservation Handler
// Paste this into Google Apps Script, deploy as Web App, then copy the URL into index.html

const SHEET_ID     = '1B6mB7vCwj8qKyxv6gMLZMlSRQ6wyA7JNRgO8NWkyTas';
const NOTIFY_EMAIL = 'iamgeorgy1@gmail.com';
const LOGO_FILE_ID = '1CyXpSJJexrSzI99Mm3rHetw2-b4iQfsX';

function doGet(e) {
  return handleRequest(e.parameter);
}

function handleRequest(params) {
  var timestamp  = new Date().toLocaleString('sl-SI', { timeZone: 'Europe/Ljubljana' });
  var p_first    = params.firstName || '';
  var p_last     = params.lastName  || '';
  var p_phone    = params.phone     || '';
  var p_email    = params.email     || '';
  var p_date     = params.date      || '';
  var p_time     = params.time      || '';
  var p_guests   = params.guests    || '';
  var p_notes    = params.notes     || '';
  var p_fullname = (p_first + ' ' + p_last).trim();

  // ── Sheet entry ──
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp','First Name','Last Name','Phone','Email','Date','Time','Guests','Notes']);
      sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#111111').setFontColor('#39FF14');
    }
    sheet.appendRow([timestamp, p_first, p_last, p_phone, p_email, p_date, p_time, p_guests, p_notes]);
  } catch (err) {
    Logger.log('Sheet error: ' + err.message);
  }

  // ── Owner notification ──
  try {
    GmailApp.sendEmail(
      NOTIFY_EMAIL,
      'New Reservation: ' + p_fullname + ' on ' + p_date + ' at ' + p_time,
      [
        'New table reservation at MAC Sportni Bar',
        '',
        'Name:   ' + p_fullname,
        'Phone:  ' + p_phone,
        'Email:  ' + p_email,
        'Date:   ' + p_date,
        'Time:   ' + p_time,
        'Guests: ' + p_guests,
        'Notes:  ' + p_notes,
        '',
        'Logged at: ' + timestamp,
        'View reservations: https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit'
      ].join('\n'),
      { name: 'MAC Bar Reservations' }
    );
  } catch (err) {
    Logger.log('Owner email error: ' + err.message);
  }

  // ── Customer confirmation ──
  try {
    if (p_email) {

      // Try to load logo from Drive for inline embedding
      var inlineImages = {};
      var logoHeader = '<div style="font-size:32px;font-weight:900;color:#fff;letter-spacing:2px;">MAC</div>'
                     + '<div style="font-size:10px;color:#39FF14;letter-spacing:4px;margin-top:4px;">SPORTNI BAR</div>';
      try {
        if (LOGO_FILE_ID !== 'YOUR_DRIVE_FILE_ID_HERE') {
          var logoBlob = DriveApp.getFileById(LOGO_FILE_ID).getBlob();
          inlineImages['maclogo'] = logoBlob;
          logoHeader = '<img src="cid:maclogo" alt="MAC Sportni Bar" style="height:90px;width:auto;display:block;margin:0 auto;">';
        }
      } catch (logoErr) {
        Logger.log('Logo load error (using text fallback): ' + logoErr.message);
      }

      var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
        + '<body style="margin:0;padding:0;background:#111111;font-family:Arial,sans-serif;">'
        + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#111111;padding:40px 20px;">'
        + '<tr><td align="center">'
        + '<table width="560" cellpadding="0" cellspacing="0" style="background:#1a1a1a;border-radius:12px;border:1px solid #3d2a1c;">'

        + '<tr><td style="background:#5A3E2B;padding:32px 40px;text-align:center;border-radius:12px 12px 0 0;border-bottom:3px solid #39FF14;">'
        + logoHeader
        + '</td></tr>'

        + '<tr><td style="padding:36px 40px 20px;">'
        + '<p style="margin:0 0 4px;font-size:12px;color:#39FF14;font-weight:700;letter-spacing:2px;text-transform:uppercase;">Rezervacija potrjena</p>'
        + '<h1 style="margin:0 0 16px;font-size:24px;font-weight:800;color:#fff;">Vaša miza je rezervirana, ' + p_first + '!</h1>'
        + '<p style="margin:0 0 0;font-size:14px;color:#a89070;line-height:1.7;">Prejeli smo vašo rezervacijo in vas z veseljem pričakujemo. Spodaj je povzetek:</p>'
        + '</td></tr>'

        + '<tr><td style="padding:20px 40px;">'
        + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#111111;border-radius:8px;border:1px solid #3d2a1c;">'
        + '<tr><td style="padding:12px 20px;border-bottom:1px solid #2a1e12;font-size:11px;color:#8a7055;text-transform:uppercase;">Ime in priimek</td>'
        + '<td style="padding:12px 20px;border-bottom:1px solid #2a1e12;font-size:14px;color:#D6C3A5;font-weight:600;text-align:right;">' + p_fullname + '</td></tr>'
        + '<tr><td style="padding:12px 20px;border-bottom:1px solid #2a1e12;font-size:11px;color:#8a7055;text-transform:uppercase;">Datum</td>'
        + '<td style="padding:12px 20px;border-bottom:1px solid #2a1e12;font-size:14px;color:#D6C3A5;font-weight:600;text-align:right;">' + p_date + '</td></tr>'
        + '<tr><td style="padding:12px 20px;border-bottom:1px solid #2a1e12;font-size:11px;color:#8a7055;text-transform:uppercase;">Ura</td>'
        + '<td style="padding:12px 20px;border-bottom:1px solid #2a1e12;font-size:14px;color:#D6C3A5;font-weight:600;text-align:right;">' + p_time + '</td></tr>'
        + '<tr><td style="padding:12px 20px;border-bottom:1px solid #2a1e12;font-size:11px;color:#8a7055;text-transform:uppercase;">Število gostov</td>'
        + '<td style="padding:12px 20px;border-bottom:1px solid #2a1e12;font-size:14px;color:#D6C3A5;font-weight:600;text-align:right;">' + p_guests + '</td></tr>'
        + '<tr><td style="padding:12px 20px;font-size:11px;color:#8a7055;text-transform:uppercase;">Opombe</td>'
        + '<td style="padding:12px 20px;font-size:14px;color:#D6C3A5;text-align:right;">' + p_notes + '</td></tr>'
        + '</table>'
        + '</td></tr>'

        + '<tr><td style="padding:24px 40px 32px;">'
        + '<p style="margin:0 0 20px;font-size:14px;color:#a89070;line-height:1.7;">Želite spremeniti rezervacijo? Pokličite nas na <strong style="color:#39FF14;">041 855 677</strong> in vse uredimo.</p>'
        + '<table cellpadding="0" cellspacing="0"><tr>'
        + '<td style="background:#39FF14;border-radius:6px;padding:12px 24px;">'
        + '<a href="https://www.instagram.com/mac.sportnibar/" style="color:#111111;text-decoration:none;font-weight:800;font-size:13px;text-transform:uppercase;">Sledite nam na Instagramu</a>'
        + '</td></tr></table>'
        + '</td></tr>'

        + '<tr><td style="background:#0d0d0d;padding:20px 40px;border-top:1px solid #2a1e12;border-radius:0 0 12px 12px;">'
        + '<p style="margin:0;font-size:11px;color:#5a4a38;line-height:1.6;">MAC Športni Bar &bull; Kašeljska cesta 66, 1260 Ljubljana - Polje &bull; 041 855 677<br>Odprto pon–sob 7:00–23:00, ned 8:00–23:00</p>'
        + '</td></tr>'

        + '</table></td></tr></table></body></html>';

      var emailOptions = {
        htmlBody: html,
        name:     'MAC Sportni Bar'
      };
      if (Object.keys(inlineImages).length > 0) {
        emailOptions.inlineImages = inlineImages;
      }

      GmailApp.sendEmail(
        p_email,
        'Rezervacija potrjena – MAC Športni Bar, ' + p_date + ' ob ' + p_time,
        'Pozdravljeni ' + p_first + ',\n\nVaša miza v MAC Športni Baru je potrjena!\n\nDatum: ' + p_date + '\nUra: ' + p_time + '\nŠtevilo gostov: ' + p_guests + '\nOpombe: ' + p_notes + '\n\nŽelite spremeniti rezervacijo? Pokličite nas na 041 855 677.\n\nZ veseljem vas pričakujemo!\nMAC Športni Bar',
        emailOptions
      );

      Logger.log('Confirmation sent to: ' + p_email);
    } else {
      Logger.log('No customer email provided, skipping confirmation.');
    }
  } catch (err) {
    Logger.log('Customer email error: ' + err.message);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
