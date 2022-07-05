using TaskWorker.Interfaces;
using Serilog;
using System;
using System.Threading.Tasks;
using EmailSender.Interfaces;
using System.Data.SqlClient;
using TaskWorker.Model;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace TaskWorker.Services
{
    public class BusinessService : IBusiness
    {
        private readonly ILogger _logger;
        private readonly IEmailSender _emailSender;
        public BusinessService(ILogger logger, IEmailSender emailSender)
        {
            _logger = logger;
            _emailSender = emailSender;
        } 
        public async Task RunBusinessTask()
        {
            _logger.Information("Task Started Here");
            var datasource = @".";//your server: DESKTOP-PC\SQLEXPRESS
            var database = "ABC"; //your database name
            var username = "sa"; //username of server to connect
            var password = "123"; //password

            //your connection string 
            string connString = @"Data Source=" + datasource + ";Initial Catalog="
                        + database + ";Persist Security Info=True;User ID=" + username + ";Password=" + password;

            //create instanace of database connection
            SqlConnection conn = new SqlConnection(connString);

            try
            {
                _logger.Information("Opening Connection to DB");
                conn.Open();

                #region First SQL 
                SqlCommand command = new SqlCommand(@"SELECT
                        [Id] make_ID
                        ,[UserId] make_name
                        ,[FirstName] model_ID
                        ,[FatherName] model_name  
                        FROM [Employees] 
                        WHERE [EmploymentTypeId]=@num", conn);
                command.Parameters.AddWithValue("@num", 1);

                List<MakeModel> makeModelList = new List<MakeModel>();

                _logger.Information("Executing MakeModel Query");

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        MakeModel makeModel = new MakeModel();
                        makeModel.make_ID = string.Format("{0}", reader["make_ID"]);
                        makeModel.make_name = string.Format("{0}", reader["make_name"]);
                        makeModel.model_ID = string.Format("{0}", reader["model_ID"]);
                        makeModel.model_name = string.Format("{0}", reader["model_name"]);
                        makeModelList.Add(makeModel);
                    }
                }
                #endregion

                #region Second SQL 
                SqlCommand command2 = new SqlCommand(@"SELECT
                       'INSURANCE_BRANCH' INSURANCE_BRANCH,
                        'INS_POLICY_KEY' INS_POLICY_KEY,
                        'INS_POLICY_REF' INS_POLICY_REF,
                        'FOB' FOB,
                        'BROKER_CODE' BROKER_CODE,
                        'CURRENCY' CURRENCY,
                        'RENEWAL_POLICY' RENEWAL_POLICY,
                        'INS_PREVPOLICY_KEY' INS_PREVPOLICY_KEY,
                        'NETWORK_KEY' NETWORK_KEY,
                        'POLICY_INDGRP' POLICY_INDGRP,
                        'HOLDER_GENDER' HOLDER_GENDER,
                        'HOLDER_FIRST_NAME_EN' HOLDER_FIRST_NAME_EN,
                        'HOLDER_MIDDLE_NAME_EN' HOLDER_MIDDLE_NAME_EN,
                        'HOLDER_LAST_NAME_EN' HOLDER_LAST_NAME_EN,
                        'HOLDER_FIRST_NAME_AR' HOLDER_FIRST_NAME_AR,
                        'HOLDER_MIDDLE_NAME_AR' HOLDER_MIDDLE_NAME_AR,
                        'HOLDER_LAST_NAME_AR' HOLDER_LAST_NAME_AR,
                        'HOLDER_DATE_OF_BIRTH' HOLDER_DATE_OF_BIRTH,
                        'HOLDER_PHONE_NUMBER' HOLDER_PHONE_NUMBER,
                        'POLICY_END' POLICY_END,
                        'END_CODE' END_CODE,
                        'EFFECTIVE' EFFECTIVE,
                        'EXPIRY' EXPIRY,
                        'NOTES' NOTES,
                        'SCHEDULE_NO' SCHEDULE_NO,
                        'PRODUCT' PRODUCT,
                        'INSURED_GENDER' INSURED_GENDER,
                        'INSURED_FIRST_NAME_EN' INSURED_FIRST_NAME_EN,
                        'INSURED_MIDDLE_NAME_EN' INSURED_MIDDLE_NAME_EN,
                        'INSURED_LAST_NAME_EN' INSURED_LAST_NAME_EN,
                        'INSURED_FIRST_NAME_AR' INSURED_FIRST_NAME_AR,
                        'INSURED_MIDDLE_NAME_AR' INSURED_MIDDLE_NAME_AR,
                        'INSURED_LAST_NAME_AR' INSURED_LAST_NAME_AR,
                        'INSURED_DATE_OF_BIRTH' INSURED_DATE_OF_BIRTH,
                        'INSURED_PHONE_NUMBER' INSURED_PHONE_NUMBER,
                        'RENEWAL_SCHEDULE' RENEWAL_SCHEDULE,
                        'SUM_INSURED' SUM_INSURED,
                        'PLATECODE' PLATECODE,
                        'PLATE' PLATE,
                        'CHASSIS' CHASSIS,
                        'MAKE' MAKE,
                        'MODEL' MODEL,
                        'USAGE' USAGE,
                        'HP' HP,
                         2022 YEAR,
                        'بطاطا' BODY,
                        'WEIGHT' WEIGHT,
                        'SEAT' SEAT,
                        'ENGINE' ENGINE,
                        'COLOR' COLOR,
                        'CC' CC,
                        'TRANS' TRANS,
                        'DOORS' DOORS,
                        'DRLIC' DRLIC,
                        'DRLICINC' DRLICINC,
                        'DRLICEXP' DRLICEXP,
                        'DEPRECIATION_SCHEDULE' DEPRECIATION_SCHEDULE,
                        'COVER' COVER,
                        'SUMINSURED' SUMINSURED,
                        'DEDUCTIBLE_CODE' DEDUCTIBLE_CODE,
                        'DEDUCTIBLE_TYPE' DEDUCTIBLE_TYPE,
                        'DEDUCTIBLE_AMNT' DEDUCTIBLE_AMNT,
                        'DEPRECIATION_SCHEDULECOVER' DEPRECIATION_SCHEDULECOVER,
                        'ALLOWED_QTTY' ALLOWED_QTTY,
                        'ALLOWED_DISTANCE_UNIT' ALLOWED_DISTANCE_UNIT,
                        'ALLOWED_DISTANCE' ALLOWED_DISTANCE,
                        'CLAUSE' CLAUSE", conn);
                //command2.Parameters.AddWithValue("@num", 1);

                List<Production> productionList = new List<Production>();

                _logger.Information("Executing Production Query");

                using (SqlDataReader reader = command2.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        Production production = new Production();
                        production.INSURANCE_BRANCH = string.Format("{0}", reader["INSURANCE_BRANCH"]);
                        production.INS_POLICY_KEY = string.Format("{0}", reader["INS_POLICY_KEY"]);
                        production.INS_POLICY_REF = string.Format("{0}", reader["INS_POLICY_REF"]);
                        production.FOB = string.Format("{0}", reader["FOB"]);
                        production.BROKER_CODE = string.Format("{0}", reader["BROKER_CODE"]);
                        production.CURRENCY = string.Format("{0}", reader["CURRENCY"]);
                        production.RENEWAL_POLICY = string.Format("{0}", reader["RENEWAL_POLICY"]);
                        production.INS_PREVPOLICY_KEY = string.Format("{0}", reader["INS_PREVPOLICY_KEY"]);
                        production.NETWORK_KEY = string.Format("{0}", reader["NETWORK_KEY"]);
                        production.POLICY_INDGRP = string.Format("{0}", reader["POLICY_INDGRP"]);
                        production.HOLDER_GENDER = string.Format("{0}", reader["HOLDER_GENDER"]);
                        production.HOLDER_FIRST_NAME_EN = string.Format("{0}", reader["HOLDER_FIRST_NAME_EN"]);
                        production.HOLDER_MIDDLE_NAME_EN = string.Format("{0}", reader["HOLDER_MIDDLE_NAME_EN"]);
                        production.HOLDER_LAST_NAME_EN = string.Format("{0}", reader["HOLDER_LAST_NAME_EN"]);
                        production.HOLDER_FIRST_NAME_AR = string.Format("{0}", reader["HOLDER_FIRST_NAME_AR"]);
                        production.HOLDER_MIDDLE_NAME_AR = string.Format("{0}", reader["HOLDER_MIDDLE_NAME_AR"]);
                        production.HOLDER_LAST_NAME_AR = string.Format("{0}", reader["HOLDER_LAST_NAME_AR"]);
                        production.HOLDER_DATE_OF_BIRTH = string.Format("{0}", reader["HOLDER_DATE_OF_BIRTH"]);
                        production.HOLDER_PHONE_NUMBER = string.Format("{0}", reader["HOLDER_PHONE_NUMBER"]);
                        production.POLICY_END = string.Format("{0}", reader["POLICY_END"]);
                        production.END_CODE = string.Format("{0}", reader["END_CODE"]);
                        production.EFFECTIVE = string.Format("{0}", reader["EFFECTIVE"]);
                        production.EXPIRY = string.Format("{0}", reader["EXPIRY"]);
                        production.NOTES = string.Format("{0}", reader["NOTES"]);
                        production.SCHEDULE_NO = string.Format("{0}", reader["SCHEDULE_NO"]);
                        production.PRODUCT = string.Format("{0}", reader["PRODUCT"]);
                        production.INSURED_GENDER = string.Format("{0}", reader["INSURED_GENDER"]);
                        production.INSURED_FIRST_NAME_EN = string.Format("{0}", reader["INSURED_FIRST_NAME_EN"]);
                        production.INSURED_MIDDLE_NAME_EN = string.Format("{0}", reader["INSURED_MIDDLE_NAME_EN"]);
                        production.INSURED_LAST_NAME_EN = string.Format("{0}", reader["INSURED_LAST_NAME_EN"]);
                        production.INSURED_FIRST_NAME_AR = string.Format("{0}", reader["INSURED_FIRST_NAME_AR"]);
                        production.INSURED_MIDDLE_NAME_AR = string.Format("{0}", reader["INSURED_MIDDLE_NAME_AR"]);
                        production.INSURED_LAST_NAME_AR = string.Format("{0}", reader["INSURED_LAST_NAME_AR"]);
                        production.INSURED_DATE_OF_BIRTH = string.Format("{0}", reader["INSURED_DATE_OF_BIRTH"]);
                        production.INSURED_PHONE_NUMBER = string.Format("{0}", reader["INSURED_PHONE_NUMBER"]);
                        production.RENEWAL_SCHEDULE = string.Format("{0}", reader["RENEWAL_SCHEDULE"]);
                        production.SUM_INSURED = string.Format("{0}", reader["SUM_INSURED"]);
                        production.PLATECODE = string.Format("{0}", reader["PLATECODE"]);
                        production.PLATE = string.Format("{0}", reader["PLATE"]);
                        production.CHASSIS = string.Format("{0}", reader["CHASSIS"]);
                        production.MAKE = string.Format("{0}", reader["MAKE"]);
                        production.MODEL = string.Format("{0}", reader["MODEL"]);
                        production.USAGE = string.Format("{0}", reader["USAGE"]);
                        production.HP = string.Format("{0}", reader["HP"]);
                        production.YEAR = string.Format("{0}", reader["YEAR"]);
                        production.BODY = string.Format("{0}", reader["BODY"]);
                        production.WEIGHT = string.Format("{0}", reader["WEIGHT"]);
                        production.SEAT = string.Format("{0}", reader["SEAT"]);
                        production.ENGINE = string.Format("{0}", reader["ENGINE"]);
                        production.COLOR = string.Format("{0}", reader["COLOR"]);
                        production.CC = string.Format("{0}", reader["CC"]);
                        production.TRANS = string.Format("{0}", reader["TRANS"]);
                        production.DOORS = string.Format("{0}", reader["DOORS"]);
                        production.DRLIC = string.Format("{0}", reader["DRLIC"]);
                        production.DRLICINC = string.Format("{0}", reader["DRLICINC"]);
                        production.DRLICEXP = string.Format("{0}", reader["DRLICEXP"]);
                        production.DEPRECIATION_SCHEDULE = string.Format("{0}", reader["DEPRECIATION_SCHEDULE"]);
                        production.COVER = string.Format("{0}", reader["COVER"]);
                        production.SUMINSURED = string.Format("{0}", reader["SUMINSURED"]);
                        production.DEDUCTIBLE_CODE = string.Format("{0}", reader["DEDUCTIBLE_CODE"]);
                        production.DEDUCTIBLE_TYPE = string.Format("{0}", reader["DEDUCTIBLE_TYPE"]);
                        production.DEDUCTIBLE_AMNT = string.Format("{0}", reader["DEDUCTIBLE_AMNT"]);
                        production.DEPRECIATION_SCHEDULECOVER = string.Format("{0}", reader["DEPRECIATION_SCHEDULECOVER"]);
                        production.ALLOWED_QTTY = string.Format("{0}", reader["ALLOWED_QTTY"]);
                        production.ALLOWED_DISTANCE_UNIT = string.Format("{0}", reader["ALLOWED_DISTANCE_UNIT"]);
                        production.ALLOWED_DISTANCE = string.Format("{0}", reader["ALLOWED_DISTANCE"]);
                        production.CLAUSE = string.Format("{0}", reader["CLAUSE"]);
                                               
                        productionList.Add(production);
                    }
                }
                #endregion

                //Convert Data to Excel and Save File 
                //TODO refactor below code and add into above loop 
                _logger.Information("Generation MakeModel Excel");
                var workbook = new XLWorkbook();
                workbook.AddWorksheet("makeModel");
                var ws = workbook.Worksheet("makeModel");
                
                int row = 1;

                #region Headers Make Model Excel
                ws.Cell("A" + row.ToString()).Value = "make_ID";
                ws.Cell("B" + row.ToString()).Value = "make_name";
                ws.Cell("C" + row.ToString()).Value = "model_ID";
                ws.Cell("D" + row.ToString()).Value = "model_name";
                #endregion 

                row = 2;
                foreach (var item in makeModelList)
                {
                    ws.Cell("A" + row.ToString()).Value = item.make_ID.ToString();
                    ws.Cell("B" + row.ToString()).Value = item.make_name.ToString();
                    ws.Cell("C" + row.ToString()).Value = item.model_ID.ToString();
                    ws.Cell("D" + row.ToString()).Value = item.model_name.ToString();
                    row++;
                }   
                _logger.Information("Saving MakeModel ExcemakeModel.model_name = l");
                var makeModelPath = "ExcelGenerated/makeModel_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
                workbook.SaveAs(makeModelPath);

                _logger.Information("Generation Production Excel");
                var workbook2 = new XLWorkbook();
                workbook2.AddWorksheet("production");
                var ws2 = workbook2.Worksheet("production");

                int row2 = 1;

                #region Headers Production Excel
                ws2.Cell("A" + row2.ToString()).Value = "INSURANCE_BRANCH";
                ws2.Cell("B" + row2.ToString()).Value = "INS_POLICY_KEY";
                ws2.Cell("C" + row2.ToString()).Value = "INS_POLICY_REF";
                ws2.Cell("D" + row2.ToString()).Value = "FOB";
                ws2.Cell("E" + row2.ToString()).Value = "BROKER_CODE";
                ws2.Cell("F" + row2.ToString()).Value = "CURRENCY";
                ws2.Cell("G" + row2.ToString()).Value = "RENEWAL_POLICY";
                ws2.Cell("H" + row2.ToString()).Value = "INS_PREVPOLICY_KEY";
                ws2.Cell("I" + row2.ToString()).Value = "NETWORK_KEY";
                ws2.Cell("J" + row2.ToString()).Value = "POLICY_INDGRP";
                ws2.Cell("K" + row2.ToString()).Value = "HOLDER_GENDER";
                ws2.Cell("L" + row2.ToString()).Value = "HOLDER_FIRST_NAME_EN";
                ws2.Cell("M" + row2.ToString()).Value = "HOLDER_MIDDLE_NAME_EN";
                ws2.Cell("N" + row2.ToString()).Value = "HOLDER_LAST_NAME_EN";
                ws2.Cell("O" + row2.ToString()).Value = "HOLDER_FIRST_NAME_AR";
                ws2.Cell("P" + row2.ToString()).Value = "HOLDER_MIDDLE_NAME_AR";
                ws2.Cell("Q" + row2.ToString()).Value = "HOLDER_LAST_NAME_AR";
                ws2.Cell("R" + row2.ToString()).Value = "HOLDER_DATE_OF_BIRTH";
                ws2.Cell("S" + row2.ToString()).Value = "HOLDER_PHONE_NUMBER";
                ws2.Cell("T" + row2.ToString()).Value = "POLICY_END";
                ws2.Cell("U" + row2.ToString()).Value = "END_CODE";
                ws2.Cell("V" + row2.ToString()).Value = "EFFECTIVE";
                ws2.Cell("W" + row2.ToString()).Value = "EXPIRY";
                ws2.Cell("X" + row2.ToString()).Value = "NOTES";
                ws2.Cell("Y" + row2.ToString()).Value = "SCHEDULE_NO";
                ws2.Cell("Z" + row2.ToString()).Value = "PRODUCT";
                ws2.Cell("AA" + row2.ToString()).Value = "INSURED_GENDER";
                ws2.Cell("AB" + row2.ToString()).Value = "INSURED_FIRST_NAME_EN";
                ws2.Cell("AC" + row2.ToString()).Value = "INSURED_MIDDLE_NAME_EN";
                ws2.Cell("AD" + row2.ToString()).Value = "INSURED_LAST_NAME_EN";
                ws2.Cell("AE" + row2.ToString()).Value = "INSURED_FIRST_NAME_AR";
                ws2.Cell("AF" + row2.ToString()).Value = "INSURED_MIDDLE_NAME_AR";
                ws2.Cell("AG" + row2.ToString()).Value = "INSURED_LAST_NAME_AR";
                ws2.Cell("AH" + row2.ToString()).Value = "INSURED_DATE_OF_BIRTH";
                ws2.Cell("AI" + row2.ToString()).Value = "INSURED_PHONE_NUMBER";
                ws2.Cell("AJ" + row2.ToString()).Value = "RENEWAL_SCHEDULE";
                ws2.Cell("AK" + row2.ToString()).Value = "SUM_INSURED";
                ws2.Cell("AL" + row2.ToString()).Value = "PLATECODE";
                ws2.Cell("AM" + row2.ToString()).Value = "PLATE";
                ws2.Cell("AN" + row2.ToString()).Value = "CHASSIS";
                ws2.Cell("AO" + row2.ToString()).Value = "MAKE";
                ws2.Cell("AP" + row2.ToString()).Value = "MODEL";
                ws2.Cell("AQ" + row2.ToString()).Value = "USAGE";
                ws2.Cell("AR" + row2.ToString()).Value = "HP";
                ws2.Cell("AS" + row2.ToString()).Value = "YEAR";
                ws2.Cell("AT" + row2.ToString()).Value = "BODY";
                ws2.Cell("AU" + row2.ToString()).Value = "WEIGHT";
                ws2.Cell("AV" + row2.ToString()).Value = "SEAT";
                ws2.Cell("AW" + row2.ToString()).Value = "ENGINE";
                ws2.Cell("AX" + row2.ToString()).Value = "COLOR";
                ws2.Cell("AY" + row2.ToString()).Value = "CC";
                ws2.Cell("AZ" + row2.ToString()).Value = "TRANS";
                ws2.Cell("BA" + row2.ToString()).Value = "DOORS";
                ws2.Cell("BB" + row2.ToString()).Value = "DRLIC";
                ws2.Cell("BC" + row2.ToString()).Value = "DRLICINC";
                ws2.Cell("BD" + row2.ToString()).Value = "DRLICEXP";
                ws2.Cell("BE" + row2.ToString()).Value = "DEPRECIATION_SCHEDULE";
                ws2.Cell("BF" + row2.ToString()).Value = "COVER";
                ws2.Cell("BG" + row2.ToString()).Value = "SUMINSURED";
                ws2.Cell("BH" + row2.ToString()).Value = "DEDUCTIBLE_CODE";
                ws2.Cell("BI" + row2.ToString()).Value = "DEDUCTIBLE_TYPE";
                ws2.Cell("BJ" + row2.ToString()).Value = "DEDUCTIBLE_AMNT";
                ws2.Cell("BK" + row2.ToString()).Value = "DEPRECIATION_SCHEDULECOVER";
                ws2.Cell("BL" + row2.ToString()).Value = "ALLOWED_QTTY";
                ws2.Cell("BM" + row2.ToString()).Value = "ALLOWED_DISTANCE_UNIT";
                ws2.Cell("BN" + row2.ToString()).Value = "ALLOWED_DISTANCE";
                ws2.Cell("BO" + row2.ToString()).Value = "CLAUSE";
                #endregion 

                row2 = 2;
                foreach (var item in productionList)
                {
                    ws2.Cell("A" + row2.ToString()).Value = item.INSURANCE_BRANCH.ToString();
                    ws2.Cell("B" + row2.ToString()).Value = item.INS_POLICY_KEY.ToString();
                    ws2.Cell("C" + row2.ToString()).Value = item.INS_POLICY_REF.ToString();
                    ws2.Cell("D" + row2.ToString()).Value = item.FOB.ToString();
                    ws2.Cell("E" + row2.ToString()).Value = item.BROKER_CODE.ToString();
                    ws2.Cell("F" + row2.ToString()).Value = item.CURRENCY.ToString();
                    ws2.Cell("G" + row2.ToString()).Value = item.RENEWAL_POLICY.ToString();
                    ws2.Cell("H" + row2.ToString()).Value = item.INS_PREVPOLICY_KEY.ToString();
                    ws2.Cell("I" + row2.ToString()).Value = item.NETWORK_KEY.ToString();
                    ws2.Cell("J" + row2.ToString()).Value = item.POLICY_INDGRP.ToString();
                    ws2.Cell("K" + row2.ToString()).Value = item.HOLDER_GENDER.ToString();
                    ws2.Cell("L" + row2.ToString()).Value = item.HOLDER_FIRST_NAME_EN.ToString();
                    ws2.Cell("M" + row2.ToString()).Value = item.HOLDER_MIDDLE_NAME_EN.ToString();
                    ws2.Cell("N" + row2.ToString()).Value = item.HOLDER_LAST_NAME_EN.ToString();
                    ws2.Cell("O" + row2.ToString()).Value = item.HOLDER_FIRST_NAME_AR.ToString();
                    ws2.Cell("P" + row2.ToString()).Value = item.HOLDER_MIDDLE_NAME_AR.ToString();
                    ws2.Cell("Q" + row2.ToString()).Value = item.HOLDER_LAST_NAME_AR.ToString();
                    ws2.Cell("R" + row2.ToString()).Value = item.HOLDER_DATE_OF_BIRTH.ToString();
                    ws2.Cell("S" + row2.ToString()).Value = item.HOLDER_PHONE_NUMBER.ToString();
                    ws2.Cell("T" + row2.ToString()).Value = item.POLICY_END.ToString();
                    ws2.Cell("U" + row2.ToString()).Value = item.END_CODE.ToString();
                    ws2.Cell("V" + row2.ToString()).Value = item.EFFECTIVE.ToString();
                    ws2.Cell("W" + row2.ToString()).Value = item.EXPIRY.ToString();
                    ws2.Cell("X" + row2.ToString()).Value = item.NOTES.ToString();
                    ws2.Cell("Y" + row2.ToString()).Value = item.SCHEDULE_NO.ToString();
                    ws2.Cell("Z" + row2.ToString()).Value = item.PRODUCT.ToString();
                    ws2.Cell("AA" + row2.ToString()).Value = item.INSURED_GENDER.ToString();
                    ws2.Cell("AB" + row2.ToString()).Value = item.INSURED_FIRST_NAME_EN.ToString();
                    ws2.Cell("AC" + row2.ToString()).Value = item.INSURED_MIDDLE_NAME_EN.ToString();
                    ws2.Cell("AD" + row2.ToString()).Value = item.INSURED_LAST_NAME_EN.ToString();
                    ws2.Cell("AE" + row2.ToString()).Value = item.INSURED_FIRST_NAME_AR.ToString();
                    ws2.Cell("AF" + row2.ToString()).Value = item.INSURED_MIDDLE_NAME_AR.ToString();
                    ws2.Cell("AG" + row2.ToString()).Value = item.INSURED_LAST_NAME_AR.ToString();
                    ws2.Cell("AH" + row2.ToString()).Value = item.INSURED_DATE_OF_BIRTH.ToString();
                    ws2.Cell("AI" + row2.ToString()).Value = item.INSURED_PHONE_NUMBER.ToString();
                    ws2.Cell("AJ" + row2.ToString()).Value = item.RENEWAL_SCHEDULE.ToString();
                    ws2.Cell("AK" + row2.ToString()).Value = item.SUM_INSURED.ToString();
                    ws2.Cell("AL" + row2.ToString()).Value = item.PLATECODE.ToString();
                    ws2.Cell("AM" + row2.ToString()).Value = item.PLATE.ToString();
                    ws2.Cell("AN" + row2.ToString()).Value = item.CHASSIS.ToString();
                    ws2.Cell("AO" + row2.ToString()).Value = item.MAKE.ToString();
                    ws2.Cell("AP" + row2.ToString()).Value = item.MODEL.ToString();
                    ws2.Cell("AQ" + row2.ToString()).Value = item.USAGE.ToString();
                    ws2.Cell("AR" + row2.ToString()).Value = item.HP.ToString();
                    ws2.Cell("AS" + row2.ToString()).Value = item.YEAR.ToString();
                    ws2.Cell("AT" + row2.ToString()).Value = item.BODY.ToString();
                    ws2.Cell("AU" + row2.ToString()).Value = item.WEIGHT.ToString();
                    ws2.Cell("AV" + row2.ToString()).Value = item.SEAT.ToString();
                    ws2.Cell("AW" + row2.ToString()).Value = item.ENGINE.ToString();
                    ws2.Cell("AX" + row2.ToString()).Value = item.COLOR.ToString();
                    ws2.Cell("AY" + row2.ToString()).Value = item.CC.ToString();
                    ws2.Cell("AZ" + row2.ToString()).Value = item.TRANS.ToString();
                    ws2.Cell("BA" + row2.ToString()).Value = item.DOORS.ToString();
                    ws2.Cell("BB" + row2.ToString()).Value = item.DRLIC.ToString();
                    ws2.Cell("BC" + row2.ToString()).Value = item.DRLICINC.ToString();
                    ws2.Cell("BD" + row2.ToString()).Value = item.DRLICEXP.ToString();
                    ws2.Cell("BE" + row2.ToString()).Value = item.DEPRECIATION_SCHEDULE.ToString();
                    ws2.Cell("BF" + row2.ToString()).Value = item.COVER.ToString();
                    ws2.Cell("BG" + row2.ToString()).Value = item.SUMINSURED.ToString();
                    ws2.Cell("BH" + row2.ToString()).Value = item.DEDUCTIBLE_CODE.ToString();
                    ws2.Cell("BI" + row2.ToString()).Value = item.DEDUCTIBLE_TYPE.ToString();
                    ws2.Cell("BJ" + row2.ToString()).Value = item.DEDUCTIBLE_AMNT.ToString();
                    ws2.Cell("BK" + row2.ToString()).Value = item.DEPRECIATION_SCHEDULECOVER.ToString();
                    ws2.Cell("BL" + row2.ToString()).Value = item.ALLOWED_QTTY.ToString();
                    ws2.Cell("BM" + row2.ToString()).Value = item.ALLOWED_DISTANCE_UNIT.ToString();
                    ws2.Cell("BN" + row2.ToString()).Value = item.ALLOWED_DISTANCE.ToString();
                    ws2.Cell("BO" + row2.ToString()).Value = item.CLAUSE.ToString();

                    row2++;
                }

                _logger.Information("Saving Production Excel");
                var productionPath = "ExcelGenerated/production_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
                workbook2.SaveAs(productionPath);

                _logger.Information("Filling Email");

                #region Specify TO emails
                List<string> ToEmails = new List<string>(); //can be read from DB or File 
                ToEmails.Add("test@test.com");
                ToEmails.Add("test2@test.com");
                #endregion

                #region Email Body Attributes
                List<string> bodyAttributes = new List<string>();
                bodyAttributes.Add("Sir/Madam"); // {{0}} can be read from DB but must be ordered
                bodyAttributes.Add("Make-Model"); // {{1}}
                bodyAttributes.Add("Production"); // {{2}}
                #endregion

                #region attachment paths
                List<string> attachmentPath = new List<string>();
                attachmentPath.Add(makeModelPath);
                attachmentPath.Add(productionPath);
                #endregion

                var mailFilled = _emailSender.FillMessage(null, ToEmails, new List<string>(), null, "readFromJson", bodyAttributes, attachmentPath);

                _logger.Information("Sending Email");
                _emailSender.SendEmail(mailFilled, "Console App Task"); //this call is not awaited
            }
            catch (Exception exp) {
                _logger.Error($"Error: {exp}");
            }
            finally { 
                conn.Close();
            }

            _logger.Information("TASK Ended");
        }
    }
}
