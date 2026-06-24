import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputPath = new URL("../seed_data/ceha-aes-2026.xlsx", import.meta.url).pathname;

const infoRows = [
  ["organization_name", "California Environmental Health Association", "Public organizer name"],
  ["organization_slug", "ceha", "Lowercase URL segment; letters, numbers, and hyphens only"],
  ["event_name", "2026 Annual Educational Symposium", "Public event name"],
  ["event_slug", "aes-2026", "Lowercase URL segment; letters, numbers, and hyphens only"],
  ["start_date", new Date("2026-04-20T12:00:00Z"), "First event date"],
  ["end_date", new Date("2026-04-24T12:00:00Z"), "Last event date"],
  ["theme", "Environmental health education, collaboration, and professional development", "Short event theme or tagline"],
  ["venue_name", "Long Beach conference venue", "Replace with the official hotel or conference center name"],
  ["venue_address", "Long Beach, California", "Full street address when available"],
  ["travel_info", "Check the official CEHA website for hotel, parking, and arrival information.", "Plain text; line breaks are supported"],
  ["local_activities", "The program encourages attendees to enjoy Long Beach nightlife on Wednesday evening.", "Nearby dining, attractions, and attendee suggestions"],
  ["website_url", "https://www.ceha.org/", "Complete URL beginning with https://"],
  ["timezone", "America/Los_Angeles", "IANA timezone used to determine past sessions"],
  ["primary_color", "#173B57", "Six-digit hex color"],
  ["secondary_color", "#E7F1F4", "Six-digit hex color"],
  ["accent_color", "#E39A3B", "Six-digit hex color"],
  ["logo_url", "", "Optional complete URL to a hosted logo image"],
  ["support_contact", "Visit the registration desk in the 2nd Floor Foyer.", "Who attendees should contact for help"],
  ["welcome_text", "Use this agenda to find sessions, speakers, rooms, and conference activities. Program details remain subject to change.", "Short welcome message"],
];

const sessions = [
  ["MON-001","Registration","Preconference registration is open.","https://www.ceha.org/","", "2026-04-20","7:00 AM","7:30 PM","2nd Floor Foyer",null,"Conference","Registration","",true,true],
  ["MON-002","Pool Plan Check","Full-day preconference training.","https://www.ceha.org/","Kiki Cason, OCEH; Kenneth Gregory, Pentair; Libby Carruesco, CDPH; Christina Buranday, OCEH; Moderator: Juan Lorenzano","2026-04-20","8:00 AM","5:00 PM","Atlantic 1&2",null,"Preconference Training","Workshop","8 CECH",true,true],
  ["MON-003","Mobile Foods","Full-day preconference training.","https://www.ceha.org/","Monica Cardenas, LB; Gabriela Mata, LB; Mobile Food Partners, Revolution; Moderator: Mozhgan Mofidi","2026-04-20","8:00 AM","5:00 PM","Pacific 1",null,"Preconference Training","Workshop","8 CECH",true,true],
  ["MON-004","Introduction to Hazard Analysis Critical Control Point (HACCP)","Full-day preconference training.","https://www.ceha.org/","Shannon Warkentin, CDPH; Johnny Velasco; Moderator: Frank O'Sullivan","2026-04-20","8:00 AM","5:00 PM","Pacific 2",null,"Preconference Training","Workshop","8 CECH",true,true],
  ["MON-005","Leadership Academy","Full-day leadership program.","https://www.ceha.org/","Ric Encarnacion; Sarah Crossman; Graciela Garcia; Linda Launer; Deepa Dillikar; Muhammed Khan; Janiah McGill","2026-04-20","8:00 AM","5:00 PM","Caribbean",null,"Preconference Training","Workshop","8 CECH",true,true],
  ["MON-006","EHSRC Program Meeting","","https://www.ceha.org/","Betty Ho","2026-04-20","10:00 AM","12:00 PM","Meetings Boardroom",null,"Meetings","Meeting","",false,true],
  ["MON-007","Lunch is on your own","","","","2026-04-20","12:00 PM","1:00 PM","",null,"Conference","Meal","",false,true],
  ["MON-008","Decorator Setup","","","","2026-04-20","12:00 PM","3:00 PM","International 1&2",null,"Exhibits","Setup","",false,true],
  ["MON-009","CEHA Board of Directors Meeting","","https://www.ceha.org/","Marissa Castillo","2026-04-20","1:00 PM","5:00 PM","Meetings Boardroom",null,"Meetings","Meeting","",false,true],
  ["MON-010","Exhibitor, Awards and Silent Auction Setup","","","","2026-04-20","3:00 PM","5:00 PM","International 1&2",null,"Exhibits","Setup","",false,true],
  ["MON-011","Exhibits Grand Opening","Early registration, prizes, hors d'oeuvres, and bar. Sponsored by Hedgerow Software.","https://www.ceha.org/","","2026-04-20","5:00 PM","7:30 PM","International 1&2",null,"Exhibits","Social","",true,true],
  ["TUE-001","Registration","","https://www.ceha.org/","","2026-04-21","7:00 AM","4:30 PM","2nd Floor Foyer",null,"Conference","Registration","",false,true],
  ["TUE-002","Exhibitor Viewing and Breakfast","","","","2026-04-21","7:00 AM","8:00 AM","International 1&2",null,"Exhibits","Breakfast","",false,true],
  ["TUE-003","Opening Ceremony and Welcome Address","","","Daryl Supernaw; Judeth Luong; Lisa Medina; Marissa Castillo; Larry Ramdin; Darryl Wong; MC: Ric Encarnacion","2026-04-21","8:00 AM","8:30 AM","International 3, 4 & 5",null,"General","Ceremony","No CECH",true,true],
  ["TUE-004","Foodborne Illness Litigation / How It Started — Better Food Safety Needed","General session keynote.","","Bill Marler; Darin Detwiler","2026-04-21","8:30 AM","10:20 AM","International 3, 4 & 5",null,"General","Keynote","1.5 CECH",true,true],
  ["TUE-005","Collaboration vs. Confrontation","General session.","","Steve Mandernach, AFDO; Chris Duggan, CRA; Shannon Warkentin, CDPH; Heather Buonomo, CCDEH Food Policy","2026-04-21","10:30 AM","11:30 AM","International 3, 4 & 5",null,"General","General Session","1.5 CECH",true,true],
  ["TUE-006","Exhibitor Viewing","","","","2026-04-21","11:30 AM","12:30 PM","International 1&2",null,"Exhibits","Exhibits","",false,true],
  ["TUE-007","Lunch and CEHA Business Meeting","Includes Leadership Program presentation and graduation.","","","2026-04-21","12:30 PM","1:50 PM","International 3, 4 & 5",480,"General","Lunch","1.5 CECH",true,true],
  ["TUE-008","Nanobubble Technology for Chemical Reduction","","","Nick Dyner, CEO Moleaer","2026-04-21","2:00 PM","2:50 PM","Atlantic 1&2",null,"Water Quality","Breakout","1 CECH",false,true],
  ["TUE-009","Introduction to Risk Based Inspection — Part 1","","","Shannon Warkentin, CDPH","2026-04-21","2:00 PM","2:50 PM","Pacific 1",null,"Communication","Breakout","1 CECH",false,true],
  ["TUE-010","Altro Floors and Walls: Compliance Driven Solutions","","","Larry McLaughlin, Altro USA","2026-04-21","2:00 PM","2:50 PM","Pacific 2",null,"Food","Breakout","1 CECH",false,true],
  ["TUE-011","Advanced Treatment Septic Systems: Protecting Our Water Resources","","","Megan Floyd, Sacramento County","2026-04-21","3:00 PM","3:50 PM","Atlantic 1&2",null,"Water Quality","Breakout","1 CECH",false,true],
  ["TUE-012","Introduction to Risk Based Inspection — Part 2","","","Shannon Warkentin, CDPH","2026-04-21","3:00 PM","3:50 PM","Pacific 1",null,"Communication","Breakout","1 CECH",false,true],
  ["TUE-013","Implementation of Microenterprise Home Kitchens in LA County","","","James Dragan, LA County; Roya Bagheri, CA; Lauren Wolfer, CA","2026-04-21","3:00 PM","3:50 PM","Pacific 2",null,"Food","Breakout","1 CECH",false,true],
  ["TUE-014","CEHA Past Presidents Council Meeting","","","Kiki Cason","2026-04-21","3:00 PM","5:00 PM","Boardroom",null,"Meetings","Meeting","",false,true],
  ["TUE-015","Seafood Inspection Program Overview and Updates","","","Shahara Warren, NOAA","2026-04-21","4:00 PM","4:50 PM","Atlantic 1&2",null,"Water Quality","Breakout","1 CECH",false,true],
  ["TUE-016","How to Effectively Communicate","","","Shannon Warkentin, CDPH","2026-04-21","4:00 PM","4:50 PM","Pacific 1",null,"Communication","Breakout","1 CECH",false,true],
  ["TUE-017","The Future of Food Safety: Emerging Trends and Innovations","","","Todd Frantz, Steritech","2026-04-21","4:00 PM","4:50 PM","Pacific 2",null,"Food","Breakout","1 CECH",false,true],
  ["TUE-018","Exhibitor Viewing and Social","Prizes, hors d'oeuvres, and bar. Sponsored by HS Gov.","","","2026-04-21","5:00 PM","7:00 PM","International 1&2",null,"Exhibits","Social","",true,true],
  ["WED-001","Registration","","","","2026-04-22","7:00 AM","4:30 PM","2nd Floor Foyer",null,"Conference","Registration","",false,true],
  ["WED-002","Exhibitor Viewing and Breakfast","","","","2026-04-22","7:00 AM","8:00 AM","International 1&2",null,"Exhibits","Breakfast","",false,true],
  ["WED-003","Bridging the Gap Between Regulated and the Regulators: Compliance to Collaboration","General session.","","Parminder Dhillon, Save Mart; MC: Michele Bilodeau","2026-04-22","8:00 AM","8:50 AM","International 3, 4 & 5",null,"General","General Session","1 CECH",true,true],
  ["WED-004","NEHA Presentation TBA","General session.","","Larry Ramdin, NEHA President","2026-04-22","9:00 AM","9:50 AM","International 3, 4 & 5",null,"General","General Session","1 CECH",true,true],
  ["WED-005","A Model Curriculum for a Competency-Based Environmental Health Practicum / Internship Experience","General session.","","Amer El-Ahraf, CSU Dominguez Hills","2026-04-22","10:00 AM","10:50 AM","International 3, 4 & 5",null,"General","General Session","1 CECH",true,true],
  ["WED-006","Exhibitor Viewing and Networking","","","","2026-04-22","11:00 AM","12:30 PM","International 1&2",null,"Exhibits","Networking","",false,true],
  ["WED-007","Lunch and CEHA Awards Ceremony","Includes the CCDEH Manager of the Year Award.","","","2026-04-22","12:30 PM","1:50 PM","International 3, 4 & 5",null,"General","Lunch","",true,true],
  ["WED-008","Wearing Multiple Hats During the Eaton Fire Response","","","Rachel Janbek, Director, Pasadena","2026-04-22","2:00 PM","2:50 PM","Atlantic 1&2",null,"Emergency Preparedness","Breakout","1 CECH",false,true],
  ["WED-009","Creating Value: Getting Public Buy-In and Support for EHS","","","Larry Ramdin, NEHA President","2026-04-22","2:00 PM","2:50 PM","Pacific 1",null,"Management","Breakout","1 CECH",false,true],
  ["WED-010","Heat Pump Water Heaters 101","","","Richard Young, Frontier Energy","2026-04-22","2:00 PM","2:50 PM","Pacific 2",null,"Heat Pump Water Heater 2.0","Breakout","1 CECH",false,true],
  ["WED-011","Emergency and Outbreak Response","","","Hermie Francisco, FDA","2026-04-22","3:00 PM","3:50 PM","Atlantic 1&2",null,"Emergency Preparedness","Breakout","1 CECH",false,true],
  ["WED-012","Being Everywhere, All at Once, at the Same Time","","","Ric Encarnacion, Director, Monterey","2026-04-22","3:00 PM","3:50 PM","Pacific 1",null,"Management","Breakout","1 CECH",false,true],
  ["WED-013","Heat Pump Water Heater Sizing Guidelines and Calculator Demonstration","","","Amin Delagah, TRC; Maya Gantley","2026-04-22","3:00 PM","3:50 PM","Pacific 2",null,"Heat Pump Water Heater 2.0","Breakout","1 CECH",false,true],
  ["WED-014","Rebuilding After the January 2025 Fires","","","Scott Abbot, LA County","2026-04-22","4:00 PM","4:50 PM","Atlantic 1&2",null,"Emergency Preparedness","Breakout","1 CECH",false,true],
  ["WED-015","Recruit, Retain, Renew: The Lifecycle of an EHS Professional","","","Freddie Agyin, Director, Vernon; April Meneghetti, Director, Yolo","2026-04-22","4:00 PM","4:50 PM","Pacific 1",null,"Management","Breakout","1 CECH",false,true],
  ["WED-016","Heat Pump Water Heater Applications in the Field and Panel Discussion","","","Richard Young; Amin Delagah; Maya Gantley, West Monroe","2026-04-22","4:00 PM","4:50 PM","Pacific 2",null,"Heat Pump Water Heater 2.0","Breakout","1 CECH",false,true],
  ["WED-017","Enjoy the Long Beach Night Life","","","","2026-04-22","5:00 PM","11:00 PM","Long Beach",null,"Social","Local Activity","",true,true],
  ["THU-001","Registration","","","","2026-04-23","7:00 AM","2:00 PM","2nd Floor Foyer",null,"Conference","Registration","",false,true],
  ["THU-002","Breakfast","","","","2026-04-23","7:00 AM","8:00 AM","Room TBA",null,"Conference","Breakfast","",false,true],
  ["THU-003","Evolving Cannabis Regulation in Los Angeles County","General session.","","Charlene Contreras, LA County; MC: Linda Launer","2026-04-23","8:00 AM","8:50 AM","International 3, 4 & 5",null,"General","General Session","1 CECH",true,true],
  ["THU-004","Industry Food Safety and the Importance of Accurate Temperature Measurements in a HACCP Program","General session.","","Jason Horn, In-N-Out; Tom Fisher, ThermoWorks; Brian Packer, ThermoWorks","2026-04-23","9:00 AM","9:50 AM","International 3, 4 & 5",480,"General","General Session","1 CECH",true,true],
  ["THU-005","Practical AI Recipes for EH Professionals","General session.","","Darryl Booth, Hedgerow Software Co.","2026-04-23","10:00 AM","10:50 AM","International 3, 4 & 5",null,"General","General Session","1 CECH",true,true],
  ["THU-006","AQMD Presentation","General session.","","Scott Epstein","2026-04-23","11:00 AM","11:50 AM","International 3, 4 & 5",null,"General","General Session","1 CECH",true,true],
  ["THU-007","Lunch and CEHA Board Installation","","","","2026-04-23","12:00 PM","1:00 PM","International 3, 4 & 5",480,"General","Lunch","",true,true],
  ["THU-008","Lithium Battery Hazard: Do You Own a Thermal Runaway?","","","David Duncan, LA Fire","2026-04-23","1:00 PM","1:50 PM","Atlantic 1&2",null,"HazMat","Breakout","1 CECH",false,true],
  ["THU-009","Air Curtains in Retail: Enhancing Food Safety and Environmental Control","","","Anday Ross, Mars Air Doors; Leslie Gerritsen, Mars Air Doors","2026-04-23","1:00 PM","1:50 PM","Pacific 1",null,"Vector Control","Breakout","1 CECH",false,true],
  ["THU-010","Body Art Inspections","","","Dean Freed; Amy Marchen; Jennifer Nguyen","2026-04-23","1:00 PM","1:50 PM","Pacific 2",null,"Body Art","Breakout","1 CECH",false,true],
  ["THU-011","CCDEH Region 4 Meeting","","","","2026-04-23","9:00 AM","3:00 PM","Boardroom",null,"Meetings","Meeting","",false,true],
  ["THU-012","Environmental Health Response to Moss Landing BESS Thermal Runaway","","","Ric Encarnacion, Director, Monterey","2026-04-23","2:00 PM","2:50 PM","Atlantic 1&2",null,"HazMat","Breakout","1 CECH",false,true],
  ["THU-013","Rodent Control in California — Challenge Accepted","","","Frank Giannico, Clark Pest Control","2026-04-23","2:00 PM","2:50 PM","Pacific 1",null,"Vector Control","Breakout","1 CECH",false,true],
  ["THU-014","Six Workplace Safety Controls for Body Piercing Facilities","","","John Johnson, OSHA","2026-04-23","2:00 PM","2:50 PM","Pacific 2",null,"Body Art","Breakout","1 CECH",false,true],
  ["THU-015","Ash to Action: Heal the Bay's Post-Fire Response","Plenary session.","","Annelisa Moe, Heal the Bay","2026-04-23","3:10 PM","4:00 PM","International 3, 4 & 5",null,"General","Plenary","1 CECH",true,true],
  ["THU-016","AES Concludes","Thank you for your support. See you next year in Sacramento at the 2027 AES.","","","2026-04-23","4:00 PM","4:15 PM","International 3, 4 & 5",null,"Conference","Closing","",true,true],
  ["FRI-001","2026–2027 CEHA Board of Directors Meeting","","","","2026-04-24","9:00 AM","5:00 PM","Boardroom",null,"Meetings","Meeting","",false,true],
];

const headers = [
  "session_number","session_name","long_description","resources_url","presenters",
  "date","start_time","end_time","room","capacity","track","session_type","cech",
  "featured","visible",
];
const scheduleRows = sessions.map((row) => {
  const converted = [...row];
  converted[5] = new Date(`${row[5]}T12:00:00Z`);
  for (const index of [6, 7]) {
    const match = row[index].match(/^(\d{1,2}):(\d{2}) (AM|PM)$/);
    let hour = Number(match[1]) % 12;
    if (match[3] === "PM") hour += 12;
    converted[index] = (hour * 60 + Number(match[2])) / (24 * 60);
  }
  return converted;
});

const workbook = Workbook.create();
const instructions = workbook.worksheets.add("Read Me");
const info = workbook.worksheets.add("Conference Info");
const schedule = workbook.worksheets.add("Sessions");
const notes = workbook.worksheets.add("Source Notes");

instructions.showGridLines = false;
instructions.getRange("A1:H2").merge();
instructions.getRange("A1").values = [["Conference Agenda Workbook"]];
instructions.getRange("A1:H2").format = {
  fill: "#173B57",
  font: { bold: true, color: "#FFFFFF", size: 22 },
  verticalAlignment: "center",
};
instructions.getRange("A4:H4").merge();
instructions.getRange("A4").values = [["How to update the attendee experience"]];
instructions.getRange("A4:H4").format = { fill: "#E7F1F4", font: { bold: true, color: "#173B57", size: 14 } };
instructions.getRange("A6:H10").values = [
  ["1", "Edit Conference Info", "Change values in column B. Keep field names in column A unchanged.", "", "", "", "", ""],
  ["2", "Edit Sessions", "Add, remove, or update rows. Keep the header row and required columns unchanged.", "", "", "", "", ""],
  ["3", "Check URLs and dates", "Use complete https:// URLs. Dates and times must be real Excel date/time values.", "", "", "", "", ""],
  ["4", "Save as .xlsx", "Do not rename the Conference Info or Sessions worksheets.", "", "", "", "", ""],
  ["5", "Upload in Admin", "A valid upload replaces the live agenda and automatically preserves the prior revision.", "", "", "", "", ""],
];
for (let row = 6; row <= 10; row += 1) {
  instructions.getRange(`C${row}:H${row}`).merge();
}
instructions.getRange("A6:A10").format = { fill: "#E39A3B", font: { bold: true, color: "#173B57", size: 14 }, horizontalAlignment: "center" };
instructions.getRange("B6:B10").format = { font: { bold: true, color: "#173B57" } };
instructions.getRange("A6:H10").format.wrapText = true;
instructions.getRange("A6:H10").format.rowHeight = 46;
instructions.getRange("A12:H14").merge();
instructions.getRange("A12").values = [["Required session fields: session_name, date, start_time, and end_time. The attendee app hides completed sessions by default, supports keyword search, and organizes the agenda by event day. Use semicolons between multiple presenters."]];
instructions.getRange("A12:H14").format = { fill: "#F6F2E9", font: { color: "#40515C" }, wrapText: true, verticalAlignment: "center" };
instructions.getRange("A:H").format.columnWidth = 14;
instructions.getRange("A:A").format.columnWidth = 6;
instructions.getRange("B:B").format.columnWidth = 24;
instructions.getRange("C:H").format.columnWidth = 14;
instructions.freezePanes.freezeRows(4);

info.showGridLines = false;
info.getRange("A1:C1").values = [["field", "value", "guidance"]];
info.getRange(`A2:C${infoRows.length + 1}`).values = infoRows;
info.getRange("A1:C1").format = { fill: "#173B57", font: { bold: true, color: "#FFFFFF" } };
info.getRange(`A2:A${infoRows.length + 1}`).format = { fill: "#E7F1F4", font: { bold: true, color: "#173B57" } };
info.getRange(`A2:C${infoRows.length + 1}`).format.wrapText = true;
info.getRange(`B6:B7`).format.numberFormat = "yyyy-mm-dd";
info.getRange("A:A").format.columnWidth = 25;
info.getRange("B:B").format.columnWidth = 54;
info.getRange("C:C").format.columnWidth = 48;
info.freezePanes.freezeRows(1);
info.tables.add(`A1:C${infoRows.length + 1}`, true, "ConferenceInfoTable").style = "TableStyleMedium2";

schedule.showGridLines = false;
schedule.getRange("A1:O1").values = [headers];
schedule.getRange(`A2:O${sessions.length + 1}`).values = scheduleRows;
schedule.getRange("A1:O1").format = { fill: "#173B57", font: { bold: true, color: "#FFFFFF" }, wrapText: true };
schedule.getRange(`A2:O${sessions.length + 1}`).format = { verticalAlignment: "top", wrapText: true };
schedule.getRange(`F2:F${sessions.length + 1}`).format.numberFormat = "yyyy-mm-dd";
schedule.getRange(`G2:H${sessions.length + 1}`).format.numberFormat = "h:mm AM/PM";
schedule.getRange(`N2:O${sessions.length + 200}`).dataValidation = { rule: { type: "list", values: ["TRUE", "FALSE"] } };
const widths = [15,36,45,30,43,13,13,13,24,11,24,18,12,10,10];
for (let index = 0; index < widths.length; index += 1) {
  schedule.getRangeByIndexes(0, index, sessions.length + 1, 1).format.columnWidth = widths[index];
}
schedule.freezePanes.freezeRows(1);
schedule.freezePanes.freezeColumns(2);
schedule.tables.add(`A1:O${sessions.length + 1}`, true, "SessionsTable").style = "TableStyleMedium2";

notes.showGridLines = false;
notes.getRange("A1:D1").values = [["source", "location", "normalization", "reason"]];
notes.getRange("A2:D7").values = [
  ["2026_aes_program_-_post_04.18.2026.pdf", "Page 3", "10:00 am – 10:50pm → 10:00 AM – 10:50 AM", "Daytime sequence makes PM an obvious source typo."],
  ["2026_aes_program_-_post_04.18.2026.pdf", "Page 3", "11:00 pm – 12:30 pm → 11:00 AM – 12:30 PM", "Placed between morning general sessions and lunch."],
  ["2026_aes_program_-_post_04.18.2026.pdf", "Page 4", "7:00 am – 2:00 mm → 7:00 AM – 2:00 PM", "The end marker is an obvious typo."],
  ["2026_aes_program_-_post_04.18.2026.pdf", "Page 4", "12:00 am – 1:00 pm → 12:00 PM – 1:00 PM", "Lunch is a midday activity."],
  ["2026_aes_program_-_post_04.18.2026.pdf", "Page 4", "11:00 am 11:50 am → 11:00 AM – 11:50 AM", "Added the missing separator."],
  ["2026_aes_program_-_post_04.18.2026.pdf", "All pages", "Generated session numbers such as TUE-008", "The source agenda does not include session numbers."],
];
notes.getRange("A1:D1").format = { fill: "#173B57", font: { bold: true, color: "#FFFFFF" } };
notes.getRange("A2:D7").format.wrapText = true;
notes.getRange("A:D").format.columnWidth = 34;
notes.getRange("B:B").format.columnWidth = 14;
notes.getRange("C:C").format.columnWidth = 48;
notes.getRange("D:D").format.columnWidth = 48;
notes.freezePanes.freezeRows(1);

await fs.mkdir(new URL("../seed_data/", import.meta.url), { recursive: true });
const preview = await workbook.render({ sheetName: "Read Me", range: "A1:H14", scale: 1.5, format: "png" });
await fs.writeFile(new URL("../seed_data/ceha-aes-2026-preview.png", import.meta.url), new Uint8Array(await preview.arrayBuffer()));
const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputPath);

const check = await workbook.inspect({
  kind: "table",
  range: "Sessions!A1:O8",
  include: "values,formulas",
  tableMaxRows: 8,
  tableMaxCols: 15,
});
console.log(check.ndjson);
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "formula error scan",
});
console.log(errors.ndjson);
console.log(outputPath);
