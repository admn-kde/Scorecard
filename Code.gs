//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//       Author:  Raymond Blocher
//        Email:  rblocher@kingsbowlamerica.com
// Date Updated:  7/30/2018
//         File:  Code.gs
//  Description:  Code file for the Scorecard fanciness in Google sheets
//////////////////////////////////////////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//    Function Name: MakeScorecardsFromTemplate
//      Description: Generates a new spreadsheet for each store and copies over the protections set up 
//       Parameters: None
//          Returns: None
//         Comments: Updated for August 2018 Scorecard
//////////////////////////////////////////////////////////////////////////////////////////////////////////////

function MakeScorecardsFromTemplate() 
{
  //Set up the string for the title of the New Scorecards
  //TODO: Alter to make the Month/Year completely automated
  var TitleP1 = 'Kings ';
  var TitleP2 = ' Daily Scorecard';
  var Month = ' August ';
  var Year = '2018';
  
  //TODO: If a new stre gets added, it will need to be added here
  var LincolnPark = 'Lincoln Park';
  var LincolnParkSheetTitle = TitleP1 + LincolnPark + TitleP2 + Month + Year;
  var Rosemont = 'Rosemont';
  var RosemontSheetTitle = TitleP1 + Rosemont + TitleP2 + Month + Year;
  var Boston = 'Boston';
  var BostonSheetTitle = TitleP1 + Boston + TitleP2 + Month + Year;
  var Burlington = 'Burlington';
  var BurlingtonSheetTitle = TitleP1 + Burlington + TitleP2 + Month + Year;
  var Dedham = 'Dedham';
  var DedhamSheetTitle = TitleP1 + Dedham + TitleP2 + Month + Year;
  var Lynnfield = 'Lynnfield';
  var LynnfieldSheetTitle = TitleP1 + Lynnfield + TitleP2 + Month + Year;
  var Seaport = 'Seaport';
  var SeaportSheetTitle = TitleP1 + Seaport + TitleP2 + Month + Year;
  var Doral = 'Doral';
  var DoralSheetTitle = TitleP1 + Doral + TitleP2 + Month + Year;
  var Franklin = 'Franklin';
  var FranklinSheetTitle = TitleP1 + Franklin + TitleP2 + Month + Year;
  var NorthHills = 'North Hills';
  var NorthHillsSheetTitle = TitleP1 + NorthHills + TitleP2 + Month + Year;
  var Orlando = 'Orlando';
  var OrlandoSheetTitle = TitleP1 + Orlando + TitleP2 + Month + Year;
  
  //Current Information about the Template (remember the ID is the portion of the URL that looks like below)
  var TemplateSheetID = '1CaKH84p5O1MTuFQSHdT_BWv_g_KhIKyMtNYqnPGr5PM';
  var TemplateSheet = DriveApp.getFileById(TemplateSheetID);
  var TemplateSheetSS = SpreadsheetApp.open(TemplateSheet);
  
  //Set up the strings to hold the information for all of the Folders for all the locations
  //TODO: See if possible to get all Subfolders iteratively, and make this part a bit less hardcoded
  //      Alternately, could move these to a Google Sheet to help this be a process others could use
  var LiveFolderID = '1VTTWRcz7wpHmKOonu_ap63hCkTOQgVKo';
  
  var LocationsFolderID = '1J98xfVrJVL8m3_SpYwKNa8v1GpaEJrxs';
  
  var MidwestFolderID = '1vi_b7z5DubQTX5pZ0HZVBTiT1F02kRlt';
   var LincolnParkFolderID = '1LUgkQ-jk7i4P0jr9s_ww8AVHj63Yh_iN';
   var RosemontFolderID = '1AvAz5-USWhqKkb9LlPyxj4B59Y_eLDLD';
  
  var NortheastFolderID = '1b9xzjOwufmX_-VL0_wzZ2BVzjngQGSOh';
   var BostonFolderID = '1svU-JrGmiz1gwjbhttQVr4ngcOZ-eeqD';
   var BurlingtonFolderID = '1ClxrfVVFvCGxdwaSj88Wep3GUfO1kHoR';
   var DedhamFolderID = '1q9TlhR7CYJDKAsmGSB5siewwFdlqFGF6';
   var LynnfieldFolderID = '1r6jMyT8WrzxD7njI6SYkf98gDhPTru9Y';
   var SeaportFolderID = '1RYR_pTYlkNAgA1frBq1EVPfndgCiCVhc';
  
  var SouthFolderID = '1tgKHKw6JdDf4Q-zkcZXGaVwmpUoWgiNI';
   var DoralFolderID = '17fR_D7cSuao6s8yAukEWV1aSILTULHS9';
   var FranklinFolderID = '1N8Ohl0Q1d8ObySgMgug7ND9k81gb9cN4';
   var NorthHillsFolderID = '1Om-rTkIKW8ZCR7TnB9HH4MNz2hscVSfG';
   var OrlandoFolderID = '1djfHrYLLDW1a78aSmJ2-PyEA_QbA3YEo';
  
  //Set up the the folders by getting them from the DriveApp with the above IDs
  var LiveFolder = DriveApp.getFolderById(LiveFolderID);
   
  var LocationsFolder = DriveApp.getFolderById(LocationsFolderID);
   
  var MidwestFolder = DriveApp.getFolderById(MidwestFolderID);
   var LincolnParkFolder = DriveApp.getFolderById(LincolnParkFolderID);
   var RosemontFolder = DriveApp.getFolderById(RosemontFolderID);
    
  var NortheastFolder = DriveApp.getFolderById(NortheastFolderID);
   var BostonFolder = DriveApp.getFolderById(BostonFolderID);
   var BurlingtonFolder = DriveApp.getFolderById(BurlingtonFolderID);
   var DedhamFolder = DriveApp.getFolderById(DedhamFolderID);
   var LynnfieldFolder = DriveApp.getFolderById(LynnfieldFolderID);
   var SeaportFolder = DriveApp.getFolderById(SeaportFolderID);
    
  var SouthFolder = DriveApp.getFolderById(SouthFolderID);
   var DoralFolder = DriveApp.getFolderById(DoralFolderID);
   var FranklinFolder = DriveApp.getFolderById(FranklinFolderID);
   var NorthHillsFolder = DriveApp.getFolderById(NorthHillsFolderID);
   var OrlandoFolder = DriveApp.getFolderById(OrlandoFolderID);
   
  
  //Use the TemplateSheetID, and the the new SheetTitle
  var LincolnParkScorecard = DriveApp.getFileById(TemplateSheetID).makeCopy(LincolnParkSheetTitle, LincolnParkFolder);
  var RosemontScorecard = DriveApp.getFileById(TemplateSheetID).makeCopy(RosemontSheetTitle, RosemontFolder);
  var BostonScorecard = DriveApp.getFileById(TemplateSheetID).makeCopy(BostonSheetTitle, BostonFolder);
  var BurlingtonScorecard = DriveApp.getFileById(TemplateSheetID).makeCopy(BurlingtonSheetTitle, BurlingtonFolder);
  var DedhamScorecard = DriveApp.getFileById(TemplateSheetID).makeCopy(DedhamSheetTitle, DedhamFolder);
  var LynnfieldScorecard = DriveApp.getFileById(TemplateSheetID).makeCopy(LynnfieldSheetTitle, LynnfieldFolder);
  var SeaportScorecard = DriveApp.getFileById(TemplateSheetID).makeCopy(SeaportSheetTitle, SeaportFolder);
  var DoralScorecard = DriveApp.getFileById(TemplateSheetID).makeCopy(DoralSheetTitle, DoralFolder);
  var FranklinScorecard = DriveApp.getFileById(TemplateSheetID).makeCopy(FranklinSheetTitle, FranklinFolder);
  var NorthHillsScorecard = DriveApp.getFileById(TemplateSheetID).makeCopy(NorthHillsSheetTitle, NorthHillsFolder);
  var OrlandoScorecard = DriveApp.getFileById(TemplateSheetID).makeCopy(OrlandoSheetTitle, OrlandoFolder);
  
  //This section is not currently necessary as all these accounts are automatically editors of the documents
  //in their respective folders
  LincolnParkScorecard.addEditor('lincolnpark.kde@gmail.com');
  RosemontScorecard.addEditor('rosemont.kde@gmail.com');
  BostonScorecard.addEditor('boston.kde@gmail.com');
  BurlingtonScorecard.addEditor('burlington.kde@gmail.com');
  DedhamScorecard.addEditor('dedham.kde@gmail.com');
  LynnfieldScorecard.addEditor('lynnfield.kde@gmail.com');
  SeaportScorecard.addEditor('seaport.kde@gmail.com');
  DoralScorecard.addEditor('doral.kde@gmail.com');
  FranklinScorecard.addEditor('franklin.kde@gmail.com');
  NorthHillsScorecard.addEditor('northhills.kde@gmail.com');
  OrlandoScorecard.addEditor('orlando.kde@gmail.com');  
  
  
  //THIS ENTIRE BLOCK IS NOT needed
//  var LincolnParkScorecardSS = SpreadsheetApp.open(LincolnParkScorecard);
//  var RosemontScorecardSS = SpreadsheetApp.open(RosemontScorecard);
//  var BostonScorecardSS = SpreadsheetApp.open(BostonScorecard);
//  var BurlingtonScorecardSS = SpreadsheetApp.open(BurlingtonScorecard);
//  var DedhamScorecardSS = SpreadsheetApp.open(DedhamScorecard);
//  var LynnfieldScorecardSS = SpreadsheetApp.open(LynnfieldScorecard);
//  var SeaportScorecardSS = SpreadsheetApp.open(SeaportScorecard);
//  var DoralScorecardSS = SpreadsheetApp.open(DoralScorecard);
//  var FranklinScorecardSS = SpreadsheetApp.open(FranklinScorecard);
//  var NorthHillsScorecardSS = SpreadsheetApp.open(NorthHillsScorecard);
//  var OrlandoScorecardSS = SpreadsheetApp.open(OrlandoScorecard);
//  
//  CopySpreadsheetProtections(TemplateSheetSS, LincolnParkScorecardSS);
//  CopySpreadsheetProtections(TemplateSheetSS, RosemontScorecardSS);
//  CopySpreadsheetProtections(TemplateSheetSS, BostonScorecardSS);
//  CopySpreadsheetProtections(TemplateSheetSS, BurlingtonScorecardSS);
//  CopySpreadsheetProtections(TemplateSheetSS, DedhamScorecardSS);
//  CopySpreadsheetProtections(TemplateSheetSS, LynnfieldScorecardSS);
//  CopySpreadsheetProtections(TemplateSheetSS, SeaportScorecardSS);
//  CopySpreadsheetProtections(TemplateSheetSS, DoralScorecardSS);
//  CopySpreadsheetProtections(TemplateSheetSS, FranklinScorecardSS);
//  CopySpreadsheetProtections(TemplateSheetSS, NorthHillsScorecardSS);
//  CopySpreadsheetProtections(TemplateSheetSS, OrlandoScorecardSS);
    
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//    Function Name: CopySpreadsheetProtections
//      Description: Will go through each sheet on the scorecard and Copy over the protections
//       Parameters: t - Template - a spreadsheet
//                   s - Scorecard - a spreadsheet
//          Returns: none
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//T for Template, S for Scorecard //These should be the Whole Spreadsheet
function CopySpreadsheetProtections(t, s)
{
  //set up the sheet variables for all the sheets in the Scorecard for the template //these ones are individual sheets within the Spreadsheet
  var tHMR = t.getSheetByName('HMR');
  var tCashout = t.getSheetByName('Cashout');
  var tDSR = t.getSheetByName('Daily Sales Report');
  var tScorecard = t.getSheetByName('Scorecard');
  var tNightlies = t.getSheetByName('Nightlies');
  var tGoalEntry = t.getSheetByName('Goal Entry');
  var tProjections = t.getSheetByName('Projections');
  var tPreMeal = t.getSheetByName('Pre-Meal');
  var tManagerSchedule = t.getSheetByName('Manager Schedule');
  var tWATipOuts = t.getSheetByName('WA TIPOUTS');
  var tWATipOutsPrintOut = t.getSheetByName('WA TIPOUTS PRINTOUT');
  var tMonthlySetUp = t.getSheetByName('Monthly Setup & Instructions');
  var tWAData = t.getSheetByName('WA DATA');
  var tBehindTheScenes = t.getSheetByName('Behind the Scenes');
  var tKPI = t.getSheetByName('KPI');
  
  //set up the sheet variables for all the sheets in the Scorecard for the new Scorecard
  var sHMR = s.getSheetByName('HMR');
  var sCashout = s.getSheetByName('Cashout');
  var sDSR = s.getSheetByName('Daily Sales Report');
  var sScorecard = s.getSheetByName('Scorecard');
  var sNightlies = s.getSheetByName('Nightlies');
  var sGoalEntry = s.getSheetByName('Goal Entry');
  var sProjections = s.getSheetByName('Projections');
  var sPreMeal = s.getSheetByName('Pre-Meal');
  var sManagerSchedule = s.getSheetByName('Manager Schedule');
  var sWATipOuts = s.getSheetByName('WA TIPOUTS');
  var sWATipOutsPrintOut = s.getSheetByName('WA TIPOUTS PRINTOUT');
  var sMonthlySetUp = s.getSheetByName('Monthly Setup & Instructions');
  var sWAData = s.getSheetByName('WA DATA');
  var sBehindTheScenes = s.getSheetByName('Behind the Scenes') ;
  var sKPI = s.getSheetByName('KPI');
  
  CopySheetProtection(tHMR, sHMR);
  CopySheetProtection(tCashout, sCashout);
  CopySheetProtection(tDSR , sDSR);
  CopySheetProtection(tScorecard , sScorecard);
  CopySheetProtection(tNightlies , sNightlies);
  CopySheetProtection(tGoalEntry , sGoalEntry);
  CopySheetProtection(tProjections , sProjections);
  CopySheetProtection(tPreMeal , sPreMeal);
  CopySheetProtection(tManagerSchedule , sManagerSchedule);
  CopySheetProtection(tWATipOuts , sWATipOuts);
  CopySheetProtection(tWATipOutsPrintOut , sWATipOutsPrintOut);
  CopySheetProtection(tMonthlySetUp , sMonthlySetUp);
  CopySheetProtection(tWAData , sWAData);
  CopySheetProtection(tBehindTheScenes , sBehindTheScenes);
  CopySheetProtection(tKPI, sKPI);
    
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//    Function Name: CopySheetProtection
//      Description: Copies over the protections including Description/etc from the template to the scorecard
//       Parameters: t - Template - a single sheet within the template scorecard
//                   s - Scorecard - a single sheet within the scorecard
//          Returns:
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
function CopySheetProtection(t, s)
{
  var protections = t.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for(var i = 0; i < protections.length; i++)
  {
    var p = protections[i];
    var rangeNotation = p.getRange().getA1Notation();
    var p2 = s.getRange(rangeNotation).protect();
    p2.setDescription(p.getDescription());
    p2.setWarningOnly(p.isWarningOnly());
    if(!p.isWarningOnly())
    {
      p2.removeEditors(p2.getEditors());
      p2.addEditors(p2.getEditors());
    }
  }
  var pr = t.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  var pr2 = s.protect();

  //pr2.setDescription(pr.getDescription());
  //pr2.setWarningOnly(pr.isWarningOnly());  
  //if (!pr.isWarningOnly()) {
    //pr2.removeEditors(p2.getEditors());
  //  pr2.addEditors(pr.getEditors());
    // p2.setDomainEdit(p.canDomainEdit()); //  only if using an Apps domain 
  //}
  //var ranges = pr.getUnprotectedRanges();
  //var newRanges = [];
  //for (var i = 0; i < ranges.length; i++) {
  //  newRanges.push(s.getRange(ranges[i].getA1Notation()));
  //} 
  //pr2.setUnprotectedRanges(newRanges);

}

/*
function convSheetAndEmail(rng, email, subj)
{
  var HTML = SheetConverter.convertRange2html(rng);
  MailApp.sendEmail(email, subj, '', {htmlBody : HTML});
}

function doGet()
{
  // or Specify a range like A1:D12, etc.
  var dataRange = SpreadsheetApp.getActiveSpreadsheet().getDataRange();

  var emailUser = 'test@email.com';

  var subject = 'Test Email';

  convSheetAndEmail(dataRange, emailUser, subject);
}
*/

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//    Function Name: SendProjectionsEmail
//      Description: -
//       Parameters: -
//          Returns: -
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
function SendProjectionsEmail()
{
  // or Specify a range like A1:D12, etc.
  //var dataRange = SpreadsheetApp.getActiveSpreadsheet().getDataRange();
  //var emailUser = 'test@email.com';
  //var subject = 'Test Email';

  //var HTML = SheetConverter.convertRange2html(rng);
  //MailApp.sendEmail(email, subj, '', {htmlBody : HTML});
  
  var ScorecardID = '1CaKH84p5O1MTuFQSHdT_BWv_g_KhIKyMtNYqnPGr5PM';
  var ScorecardSheetFile = DriveApp.getFileById(ScorecardID);
  var ScorecardSpreadsheet = SpreadsheetApp.open(ScorecardSheetFile);
  var ProjectionsSheet = ScorecardSpreadsheet.getSheetByName('Projections');
  var EmailsSheet = ScorecardSpreadsheet.getSheetByName('Emails');
  
  var LastWeekSentRange = EmailsSheet.getRange(16, 1);
  var LWSValue = LastWeekSentRange.getValue();
  
  var TopRow = 52;
  var TopChartRow = 56;
  var BotRow = 65;
  var NumRowsTop = 4;
  var NumRowsBot = 3;
  var NumRowsChart = 9;
  var LeftCol;
  var NumColsChart = 8;
  var NumCols = 1;
  
  if(LWSValue == 0)
    LeftCol = 4;  
  if(LWSValue == 1)
    LeftCol = 13;
  if(LWSValue == 2)
    LeftCol = 22;
  if(LWSValue == 3)
    LeftCol = 31;
  if(LWSValue == 4)
    LeftCol = 40;
  
  if(LWSValue == 5)
    return; 
    
  var EmailSubject = ProjectionsSheet.getRange(TopRow, LeftCol).getValue();
  var EmailRecipients = 'rblocher@kingsbowlamerica.com';
  var ProjectionsContentsRangeChart = ProjectionsSheet.getRange(TopChartRow, LeftCol, NumRowsChart, NumColsChart);
  var ProjectionsContentsRangeTop = ProjectionsSheet.getRange(TopRow, LeftCol, NumRowsTop, NumCols);
  var ProjectionsContentsRangeBot = ProjectionsSheet.getRange(BotRow, LeftCol, NumRowsBot, NumCols);
  var HTMLChart = SheetConverter.convertRange2html(ProjectionsContentsRangeChart);
  var HTMLTop = SheetConverter.convertRange2html(ProjectionsContentsRangeTop);
  var HTMLBot = SheetConverter.convertRange2html(ProjectionsContentsRangeBot);
  
  //Specify the target in the <a href=" ">.
  //Then add the text that should work as a link.
  //Finally add an </a> tag to indicate where the link ends.
  var LinkHTML = '<a href=\"' +ScorecardSpreadsheet.getUrl()+ '">Scorecard Link</a>';
  var HTMLAll = HTMLChart + LinkHTML;
  
  MailApp.sendEmail(EmailRecipients, EmailSubject,' ', {htmlBody : HTMLAll}) ;
  Logger.log('Projections Email sent.');
  
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//    Function Name: SendHMREmail
//      Description: -
//       Parameters: -
//          Returns: -
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
function SendHMREmail()
{
  // or Specify a range like A1:D12, etc.
  //var dataRange = SpreadsheetApp.getActiveSpreadsheet().getDataRange();
  //var emailUser = 'test@email.com';
  //var subject = 'Test Email';

//var HTML = SheetConverter.convertRange2html(rng);
  //MailApp.sendEmail(email, subj, '', {htmlBody : HTML});
  
  var ScorecardID = '1CaKH84p5O1MTuFQSHdT_BWv_g_KhIKyMtNYqnPGr5PM';
  var ScorecardSheetFile = DriveApp.getFileById(ScorecardID);
  var ScorecardSpreadsheet = SpreadsheetApp.open(ScorecardSheetFile);
  var HMRSheet = ScorecardSpreadsheet.getSheetByName('HMR');
  var EmailsSheet = ScorecardSpreadsheet.getSheetByName('Emails');
  
  var LastWeekSentRange = EmailsSheet.getRange(16, 1);
  var LWSValue = LastWeekSentRange.getValue();
  
  var TopRow;
  var NumRows = 27;
  var LeftCol = 1;
  var NumCols = 16;
  
  if(LWSValue == 0)
    TopRow = 3;  
  if(LWSValue == 1)
    TopRow = 31;
  if(LWSValue == 2)
    TopRow = 59;
  if(LWSValue == 3)
    TopRow = 87;
  if(LWSValue == 4)
    TopRow = 115;
  
  if(LWSValue == 5)
    return; 
   
  var EmailSubject = 'Head Mechanic Report Week ' + String(LWSValue+1);
  var EmailRecipients = 'rblocher@kingsbowlamerica.com';
  var HMRSheetContentsRange = HMRSheet.getRange(TopRow, LeftCol, NumRows, NumCols);
  var HTMLHMR = SheetConverter.convertRange2html(HMRSheetContentsRange);
  
  MailApp.sendEmail(EmailRecipients, EmailSubject,' ', {htmlBody : HTMLHMR}) ;
  Logger.log('HMR Email sent.');
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//    Function Name: SendNightlyEmail
//      Description: -
//       Parameters: -
//          Returns: -
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
function SendNightlyEmail()
{
  var ScorecardID = '1CaKH84p5O1MTuFQSHdT_BWv_g_KhIKyMtNYqnPGr5PM';
  var ScorecardSheetFile = DriveApp.getFileById(ScorecardID);
  var ScorecardSpreadsheet = SpreadsheetApp.open(ScorecardSheetFile);
  
  var ScorecardSheet = ScorecardSpreadsheet.getSheetByName('Scorecard');
  var EmailNightliesSheet = ScorecardSpreadsheet.getSheetByName('Email Nightlies');
  var EmailCashoutSheet = ScorecardSpreadsheet.getSheetByName('Email Cashout');
  var EmailManagerScheduleSheet = ScorecardSpreadsheet.getSheetByName('Email Manager Schedule');
  var EmailsSheet = ScorecardSpreadsheet.getSheetByName('Emails');
  var KPISheet = ScorecardSpreadsheet.getSheetByName('KPI');
  
  var NumOfWeeks = KPISheet.getRange(4, 3).getValue();
  
  var ScorecardURL = ScorecardSheetFile.getUrl(); 
  //ScorecardSpreadsheet.getActiveSheet().getUrl();
  
  //ToDo: Get the appropriate areas of each sheet to send out
  //Get the appropriate list of emails for each Store to send the email to
  var StoreListEmail = EmailsSheet.getRange(19, 2).getValue();
  var RegionListEmail = EmailsSheet.getRange(19, 3).getValue();
  var AdminListEmail = EmailsSheet.getRange(19, 4).getValue();
  
  var EmailCurrDayRange = EmailsSheet.getRange(20,1);
  var EmailCurrDay = EmailCurrDayRange.getValue();
  
  //exit early if we're not in a 5 week month
  if(EmailCurrDay > 28 && NumofWeeks == 4)
    return;
  
  var NightliesColumnStart = EmailsSheet.getRange(20, 2).getValue();
  var NightliesRowStart = EmailsSheet.getRange(20,3).getValue();
  var CashoutEmployeeRowStart = EmailsSheet.getRange(20,4).getValue();
  var CashoutEmployeeRowEnd = EmailsSheet.getRange(20,5).getValue();
  var CashoutManagerRowStart = EmailsSheet.getRange(20,6).getValue();
  var CashoutManagerRowEnd = EmailsSheet.getRange(20,7).getValue();
  var CashoutEmployeeCount = EmailCashoutSheet.getRange(CashoutEmployeeRowStart, 9).getValue();
  var CashoutManagerCount = EmailCashoutSheet.getRange(CashoutEmployeeRowStart, 10).getValue();
  var ScorecardRowStart = EmailsSheet.getRange(20,8).getValue();
  var ScorecardRowEnd = EmailsSheet.getRange(20,9).getValue();
  var ManagerScheduleColumnStart = EmailsSheet.getRange(20,10).getValue();
  var ManagerScheduleRowStart = EmailsSheet.getRange(20, 11).getValue();
    
  var EmailSubject = EmailNightliesSheet.getRange(NightliesRowStart-1, NightliesColumnStart+1).getValue();
  var EmailRecipients = 'rblocher@kingsbowlamerica.com';//StoreListEmail + ',' + RegionListEmail + ',' + AdminListEmail;//'rblocher@kingsbowlamerica.com';
  
  var EmailLinkHTML = '<a href="' + ScorecardURL + '"> Scorecard Link</a>';
  
  var EmailCashoutEmployeeRange;
  var EmailCashoutManagerRange;
  var EmailManagerScheduleRange = EmailManagerScheduleSheet.getRange(ManagerScheduleRowStart, 
                                  ManagerScheduleColumnStart, 12, 4);
  var EmailNightliesRange = EmailNightliesSheet.getRange(NightliesRowStart, NightliesColumnStart, 18, 2);
  var EmailCashoutHeaderRange = EmailCashoutSheet.getRange(2, 1, 1, 7);
  if(CashoutEmployeeCount > 0)
     EmailCashoutEmployeeRange = EmailCashoutSheet.getRange(CashoutEmployeeRowStart, 1, CashoutEmployeeCount, 7);
  if(CashoutManagerCount > 0)
     EmailCashoutManagerRange = EmailCashoutSheet.getRange(CashoutManagerRowStart, 1, CashoutManagerCount, 7);
    
  var EmailScorecardRange = ScorecardSheet.getRange(ScorecardRowStart, 1, (ScorecardRowEnd-ScorecardRowStart), 10);
  
  var EmailCashoutEmployeeHTML;
  var EmailCashoutManagerHTML;
  var EmailManagerScheduleHTML = SheetConverter.convertRange2html(EmailManagerScheduleRange);
  var EmailNightliesHTML = SheetConverter.convertRange2html(EmailNightliesRange);
  var EmailCashoutHeaderHTML = SheetConverter.convertRange2html(EmailCashoutHeaderRange);
  if(CashoutEmployeeCount > 0)
    EmailCashoutEmployeeHTML = SheetConverter.convertRange2html(EmailCashoutEmployeeRange);
  else
    EmailCashoutEmployeeHTML = '';
  if(CashoutManagerCount > 0)
    EmailCashoutManagerHTML = SheetConverter.convertRange2html(EmailCashoutManagerRange);
  else
    EmailCashoutManagerHTML = '';
  var EmailScorecardHTML = SheetConverter.convertRange2html(EmailScorecardRange);
  
  
  
  var HTMLAll = EmailLinkHTML + EmailScorecardHTML + EmailCashoutHeaderHTML + EmailCashoutEmployeeHTML +
                EmailManagerScheduleHTML;
  //Old Format
  //EmailNightliesHTML + EmailCashoutHeaderHTML + EmailCashoutEmployeeHTML + EmailCashoutManagerHTML + 
  //              EmailScorecardHTML;
  
  MailApp.sendEmail(EmailRecipients, EmailSubject,' ', {htmlBody : HTMLAll}) ;
  //SpreadsheetApp.getUi().alert(EmailSubject + ' , ' + EmailRecipients);
  EmailCurrDayRange.setValue(EmailCurrDay+1);
  //Logger.log('Nightly Email sent.');
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//    Function Name: CashoutSettings
//      Description: -
//       Parameters: -
//          Returns: -
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
function CashoutSettings()
{
  var ScorecardID = '1CaKH84p5O1MTuFQSHdT_BWv_g_KhIKyMtNYqnPGr5PM';
  var CashoutSheetFile = DriveApp.getFileById(ScorecardID);
  var CashoutSpreadsheet = SpreadsheetApp.open(CashoutSheetFile);
  var CashoutSheet = CashoutSpreadsheet.getSheetByName('Cashout');
  
  var NightliesButtonRange = CashoutSheet.getRange(1, 5);
  var DiscountButtonRange = CashoutSheet.getRange(2, 14);
  var DateSelectionRange = CashoutSheet.getRange(1, 7);
  var PositionSelectionRange = CashoutSheet.getRange(1, 11);

  //var CurrSelectionValue = CurrSelection.getValue();
  var NightliesButtonRangeValue = NightliesButtonRange.getValue();
  var DiscountButtonRangeValue = DiscountButtonRange.getValue();
  var DateSelectionRangeValue = DateSelectionRange.getValue();
  var PositionSelectionRangeValue = PositionSelectionRange.getValue();
  
  var DateSelectionRangeValueDate = new Date(DateSelectionRangeValue);
  
  var LSSNightlies = CashoutSheet.getRange(3, 24).getValue();
  var LSSDiscount = CashoutSheet.getRange(3, 25).getValue();
  var LSSDateSelect = CashoutSheet.getRange(3, 26).getValue();
  var LSSPositionSelect = CashoutSheet.getRange(3,27).getValue();
  
  var LSSDateDate = new Date(LSSDateSelect);
  
  //if this is true then there has been no change on the sheet since the last time we checked.
  if( (LSSNightlies == NightliesButtonRangeValue) && (LSSDiscount == DiscountButtonRangeValue) && (LSSDateDate.getTime() == DateSelectionRangeValueDate.getTime()) && (LSSPositionSelect == PositionSelectionRangeValue))
    return;
    
   if((LSSNightlies != NightliesButtonRangeValue))
  {
    if(NightliesButtonRangeValue == 'Nightly')
    {
      CashoutSheet.hideColumns(7, 8);
      CashoutSheet.hideColumns(17, 13);
    }
    else
    {
      CashoutSheet.showColumns(7, 8);
      CashoutSheet.showColumns(17, 13);
      CashoutSheet.hideColumns(12, 2);
    }
    CashoutSheet.getRange(3, 24).setValue(NightliesButtonRangeValue);
  }

  if((LSSDiscount != DiscountButtonRangeValue))
  {
    if(DiscountButtonRangeValue == 'Show Discounts')
      CashoutSheet.showColumns(17, 11);
    else
      CashoutSheet.hideColumns(17, 11);
    CashoutSheet.getRange(3, 25).setValue(DiscountButtonRangeValue);
  }

    var SBStart = Number(CashoutSheet.getRange(3, 17).getValue());
    var SBEnd = Number(CashoutSheet.getRange(3, 18).getValue());
    var MStart = Number(CashoutSheet.getRange(3, 19).getValue());
    var MEnd = Number(CashoutSheet.getRange(3, 20).getValue());
    
    var OSBStart = Number(CashoutSheet.getRange(2, 17).getValue());
    var OSBEnd = Number(CashoutSheet.getRange(2, 18).getValue());
    var OMStart = Number(CashoutSheet.getRange(2, 19).getValue());
    var OMEnd = Number(CashoutSheet.getRange(2, 20).getValue());
    
    //if(DateSelectionRangeValue == 'All')
    //{
    //  CashoutSheet.showRows(5, 2238);
    //}
  if( (LSSDateDate.getTime() != DateSelectionRangeValueDate.getTime()) || (LSSPositionSelect != PositionSelectionRangeValue) )
  {
    if(PositionSelectionRangeValue == 'Servers & Bartenders')
    {
      if(DateSelectionRangeValue == 'All Days')
      {
        if(LSSDateSelect == 'All Days')
          CashoutHideAllDays(CashoutSheet);
        else
        {
          CashoutSheet.hideRows(OSBStart, (OSBEnd-OSBStart)+1);
          CashoutSheet.hideRows(OMStart, (OMEnd-OMStart)+1);
        }
        CashoutShowAllSB(CashoutSheet);
      }
      else
      {
        if(LSSDateSelect == 'All Days')
          CashoutHideAllDays(CashoutSheet);
        else
        {
          CashoutSheet.hideRows(OSBStart, (OSBEnd-OSBStart)+1);
          CashoutSheet.hideRows(OMStart, (OMEnd-OMStart)+1);
        }
        CashoutSheet.showRows(SBStart, (SBEnd-SBStart)+1);
      }
      CashoutSheet.showColumns(7, 23);
      CashoutSheet.hideColumns(12, 2);
    }
    else if(PositionSelectionRangeValue == 'Managers')
    {
      if(DateSelectionRangeValue == 'All Days')
      {
        if(LSSDateSelect == 'All Days')
          CashoutHideAllDays(CashoutSheet);
        else
        {
          CashoutSheet.hideRows(OSBStart, (OSBEnd-OSBStart)+1);
          CashoutSheet.hideRows(OMStart, (OMEnd-OMStart)+1);
        }
        CashoutShowAllManagers(CashoutSheet);
      }
      else
      {
        if(LSSDateSelect == 'All Days')
          CashoutHideAllDays(CashoutSheet);
        else
        {
          CashoutSheet.hideRows(OSBStart, (OSBEnd-OSBStart)+1);
          CashoutSheet.hideRows(OMStart, (OMEnd-OMStart)+1);
        }
        CashoutSheet.showRows(MStart, (MEnd-MStart)+1);
      }
      //CashoutSheet.hideColumns(7, 23);
      //CashoutSheet.showColumns(11, 4);
      //CashoutSheet.hideColumns(12, 2);
    }
    else
    {
      if(DateSelectionRangeValue == 'All Days')
      {
        CashoutShowAll(CashoutSheet);
      }
      else
      {
        if(LSSDateSelect == 'All Days')
          CashoutHideAllDays(CashoutSheet);
        else
        {
          CashoutSheet.hideRows(OSBStart, (OSBEnd-OSBStart)+1);
          CashoutSheet.hideRows(OMStart, (OMEnd-OMStart)+1);
        }
        CashoutSheet.showRows(SBStart, (SBEnd-SBStart)+1);
        CashoutSheet.showRows(MStart, (MEnd-MStart)+1);
      }
    }
    CashoutSheet.getRange(3, 26).setValue(DateSelectionRangeValue);
    CashoutSheet.getRange(3, 27).setValue(PositionSelectionRangeValue);
  }
  
}

function CashoutShowAll(co)
{
  //var CashoutSheet = SpreadsheetApp.getActiveSheet();
  //if(CashoutSheet.getSheetName() != "Cashout")
  //  return;
    
  var CashoutSheet = co;
  CashoutShowAllSB(co);
  CashoutShowAllManagers(co);
}

function CashoutShowAllSB(co)
{
  //var CashoutSheet = SpreadsheetApp.getActiveSheet();
  //if(CashoutSheet.getSheetName() != "Cashout")
  //  return;
    
  var CashoutSheet = co;
  CashoutSheet.showRows(5 ,39);
  CashoutSheet.showRows(69,39);
  CashoutSheet.showRows(133 ,39);
  CashoutSheet.showRows(197 ,39);
  CashoutSheet.showRows(261 ,39);
  CashoutSheet.showRows(325 ,39);
  CashoutSheet.showRows(389 ,39);
  CashoutSheet.showRows(453 ,39);
  CashoutSheet.showRows(517 ,39);
  CashoutSheet.showRows(581 ,39);
  CashoutSheet.showRows(645 ,39);
  CashoutSheet.showRows(709 ,39);
  CashoutSheet.showRows(773 ,39);
  CashoutSheet.showRows(837 ,39);
  CashoutSheet.showRows(901 ,39);
  CashoutSheet.showRows(965 ,39);
  CashoutSheet.showRows(1029 ,39);
  CashoutSheet.showRows(1093 ,39);
  CashoutSheet.showRows(1157 ,39);
  CashoutSheet.showRows(1221 ,39);
  CashoutSheet.showRows(1285 ,39);
  CashoutSheet.showRows(1349 ,39);
  CashoutSheet.showRows(1413 ,39);
  CashoutSheet.showRows(1477 ,39);
  CashoutSheet.showRows(1541 ,39);
  CashoutSheet.showRows(1605 ,39);
  CashoutSheet.showRows(1669 ,39);
  CashoutSheet.showRows(1733 ,39);
  CashoutSheet.showRows(1797 ,39);
  CashoutSheet.showRows(1861 ,39);
  CashoutSheet.showRows(1925 ,39);
  CashoutSheet.showRows(1989 ,39);
  CashoutSheet.showRows(2053 ,39);
  CashoutSheet.showRows(2117 ,39);
  CashoutSheet.showRows(2181 ,39);
}

function CashoutShowAllManagers(co)
{
  //var CashoutSheet = SpreadsheetApp.getActiveSheet();
  //if(CashoutSheet.getSheetName() != "Cashout")
  //  return;
  
  var CashoutSheet = co;
  CashoutSheet.showRows(44  ,24);
  CashoutSheet.showRows(108  ,24);
  CashoutSheet.showRows(172  ,24);
  CashoutSheet.showRows(236  ,24);
  CashoutSheet.showRows(300  ,24);
  CashoutSheet.showRows(364  ,24);
  CashoutSheet.showRows(428  ,24);
  CashoutSheet.showRows(492  ,24);
  CashoutSheet.showRows(556  ,24);
  CashoutSheet.showRows(620  ,24);
  CashoutSheet.showRows(684  ,24);
  CashoutSheet.showRows(748  ,24);
  CashoutSheet.showRows(812  ,24);
  CashoutSheet.showRows(876  ,24);
  CashoutSheet.showRows(940  ,24);
  CashoutSheet.showRows(1004  ,24);
  CashoutSheet.showRows(1068  ,24);
  CashoutSheet.showRows(1132  ,24);
  CashoutSheet.showRows(1196  ,24);
  CashoutSheet.showRows(1260  ,24);
  CashoutSheet.showRows(1324  ,24);
  CashoutSheet.showRows(1388  ,24);
  CashoutSheet.showRows(1452  ,24);
  CashoutSheet.showRows(1516  ,24);
  CashoutSheet.showRows(1580  ,24);
  CashoutSheet.showRows(1644  ,24);
  CashoutSheet.showRows(1708  ,24);
  CashoutSheet.showRows(1772  ,24);
  CashoutSheet.showRows(1836  ,24);
  CashoutSheet.showRows(1900  ,24);
  CashoutSheet.showRows(1964  ,24);
  CashoutSheet.showRows(2028  ,24);
  CashoutSheet.showRows(2092  ,24);
  CashoutSheet.showRows(2156  ,24);
  CashoutSheet.showRows(2220  ,24);
}

function CashoutHideAllSB(co)
{
  //var CashoutSheet = SpreadsheetApp.getActiveSheet();
  //if(CashoutSheet.getSheetName() != "Cashout")
  //  return;
    
  var CashoutSheet = co;
  CashoutSheet.hideRows(5 ,39);
  CashoutSheet.hideRows(69,39);
  CashoutSheet.hideRows(133 ,39);
  CashoutSheet.hideRows(197 ,39);
  CashoutSheet.hideRows(261 ,39);
  CashoutSheet.hideRows(325 ,39);
  CashoutSheet.hideRows(389 ,39);
  CashoutSheet.hideRows(453 ,39);
  CashoutSheet.hideRows(517 ,39);
  CashoutSheet.hideRows(581 ,39);
  CashoutSheet.hideRows(645 ,39);
  CashoutSheet.hideRows(709 ,39);
  CashoutSheet.hideRows(773 ,39);
  CashoutSheet.hideRows(837 ,39);
  CashoutSheet.hideRows(901 ,39);
  CashoutSheet.hideRows(965 ,39);
  CashoutSheet.hideRows(1029 ,39);
  CashoutSheet.hideRows(1093 ,39);
  CashoutSheet.hideRows(1157 ,39);
  CashoutSheet.hideRows(1221 ,39);
  CashoutSheet.hideRows(1285 ,39);
  CashoutSheet.hideRows(1349 ,39);
  CashoutSheet.hideRows(1413 ,39);
  CashoutSheet.hideRows(1477 ,39);
  CashoutSheet.hideRows(1541 ,39);
  CashoutSheet.hideRows(1605 ,39);
  CashoutSheet.hideRows(1669 ,39);
  CashoutSheet.hideRows(1733 ,39);
  CashoutSheet.hideRows(1797 ,39);
  CashoutSheet.hideRows(1861 ,39);
  CashoutSheet.hideRows(1925 ,39);
  CashoutSheet.hideRows(1989 ,39);
  CashoutSheet.hideRows(2053 ,39);
  CashoutSheet.hideRows(2117 ,39);
  CashoutSheet.hideRows(2181 ,39);
}

function CashoutHideAllManagers(co)
{
  //var CashoutSheet = SpreadsheetApp.getActiveSheet();
  //if(CashoutSheet.getSheetName() != "Cashout")
  //  return;
    
  var CashoutSheet = co;
  CashoutSheet.hideRows(44  ,24);
  CashoutSheet.hideRows(108  ,24);
  CashoutSheet.hideRows(172  ,24);
  CashoutSheet.hideRows(236  ,24);
  CashoutSheet.hideRows(300  ,24);
  CashoutSheet.hideRows(364  ,24);
  CashoutSheet.hideRows(428  ,24);
  CashoutSheet.hideRows(492  ,24);
  CashoutSheet.hideRows(556  ,24);
  CashoutSheet.hideRows(620  ,24);
  CashoutSheet.hideRows(684  ,24);
  CashoutSheet.hideRows(748  ,24);
  CashoutSheet.hideRows(812  ,24);
  CashoutSheet.hideRows(876  ,24);
  CashoutSheet.hideRows(940  ,24);
  CashoutSheet.hideRows(1004  ,24);
  CashoutSheet.hideRows(1068  ,24);
  CashoutSheet.hideRows(1132  ,24);
  CashoutSheet.hideRows(1196  ,24);
  CashoutSheet.hideRows(1260  ,24);
  CashoutSheet.hideRows(1324  ,24);
  CashoutSheet.hideRows(1388  ,24);
  CashoutSheet.hideRows(1452  ,24);
  CashoutSheet.hideRows(1516  ,24);
  CashoutSheet.hideRows(1580  ,24);
  CashoutSheet.hideRows(1644  ,24);
  CashoutSheet.hideRows(1708  ,24);
  CashoutSheet.hideRows(1772  ,24);
  CashoutSheet.hideRows(1836  ,24);
  CashoutSheet.hideRows(1900  ,24);
  CashoutSheet.hideRows(1964  ,24);
  CashoutSheet.hideRows(2028  ,24);
  CashoutSheet.hideRows(2092  ,24);
  CashoutSheet.hideRows(2156  ,24);
  CashoutSheet.hideRows(2220  ,24);
}

function CashoutHideAllDays(co)
{
  //var CashoutSheet = SpreadsheetApp.getActiveSheet();
  //if(CashoutSheet.getSheetName() != "Cashout")
  //  return;
    
  var CashoutSheet = co;
  //Hide all of the Server & Bartender Areas
  CashoutHideAllSB(CashoutSheet);
  CashoutHideAllManagers(CashoutSheet);
}

