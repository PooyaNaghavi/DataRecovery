using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace DataRecovery
{
    class Program
    {
        static void Main(string[] args)
        {
          

            //Execute(24, 31, "k");
            Execute(22, 24, "t");
            //Execute(22, 24, "t");

        }
        public static void ConvertToJson(List<Data> data, int i, string month)
        {
            string json = JsonConvert.SerializeObject(data.ToArray());
            
            //write string o file
            System.IO.File.WriteAllText(@"C:\Users\pooya\Desktop\path"+ i + month +".txt", json);
        }
        public static void Execute(int first, int last, string month)
        {
            for (int k = first; k <= last; k++)
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\pooya\Desktop\Send\" + k + month+".xls");

                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                
                int counter = 0;

                List<Data> data = new List<Data>();
                for (int i = 1; i <= xlRange.Rows.Count; i++)
                {
                    for (int j = 1; j <= xlRange.Columns.Count; j++)
                    {
                        if (j == 1)
                            Console.Write("\r\n");
                        
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            string command = (string)xlRange.Cells[i, j].Value2.ToString();
                            string partialConmmand = "";
                            if (command.Length >= 20)
                                partialConmmand = command.Substring(0, 20);
                            if (j == 2 && !partialConmmand.Equals("") && !xlRange.Cells[i, 3].Value2.ToString().Equals("10009099099099"))
                            {
                                if (partialConmmand.Equals("با سلام\nبه رسا خوش آ"))
                                {
                                    counter++;
                                    Console.WriteLine("Patient Registred");
                                    data = PatientRegistred(data, (string)xlRange.Cells[i, 2].Value2.ToString(),
                                        (string)xlRange.Cells[i, 4].Value2.ToString(), (string)xlRange.Cells[i, 5].Value2.ToString());

                                }
                                else if (partialConmmand.Equals("با سلام. از این پس ا"))
                                {
                                    counter++;
                                    Console.WriteLine("Call Blocked Without Docter Name");
                                    data = CallBockWithoutName(data, (string)xlRange.Cells[i, 2].Value2.ToString(),
                                        (string)xlRange.Cells[i, 4].Value2.ToString(), (string)xlRange.Cells[i, 5].Value2.ToString());
                                }
                                else if (partialConmmand.Equals("سامانه رسا\nاز تماس ش"))
                                {
                                    counter++;
                                    Console.WriteLine("Feedback After Promotional Call");
                                    data = FeedbackAfterPromotionalCall(data, (string)xlRange.Cells[i, 2].Value2.ToString(),
                                        (string)xlRange.Cells[i, 4].Value2.ToString(), (string)xlRange.Cells[i, 5].Value2.ToString());
                                }
                                else if (partialConmmand.Equals("سامانه رسا\nمدت تماس:"))
                                {
                                    counter++;
                                    Console.WriteLine("FeedBack");
                                    data = Feedback(data, (string)xlRange.Cells[i, 2].Value2.ToString(),
                                        (string)xlRange.Cells[i, 4].Value2.ToString(), (string)xlRange.Cells[i, 5].Value2.ToString());
                                }
                                else if (partialConmmand.Equals("سامانه رسا\nبا سلام. "))
                                {
                                    counter++;
                                    Console.WriteLine("Call Blocked With Docter Name");
                                    data = CallBlockWithName(data, (string)xlRange.Cells[i, 2].Value2.ToString(),
                                        (string)xlRange.Cells[i, 4].Value2.ToString(), (string)xlRange.Cells[i, 5].Value2.ToString());
                                }
                                else if (partialConmmand.Equals("سامانه رسا\nکارت شما "))
                                {
                                    counter++;
                                    Console.WriteLine("One Time Charge Notification");
                                    data = OneTimeChargeNotification(data, (string)xlRange.Cells[i, 2].Value2.ToString(),
                                        (string)xlRange.Cells[i, 4].Value2.ToString(), (string)xlRange.Cells[i, 5].Value2.ToString());
                                }
                                else if (partialConmmand.Equals("پزشک محترم\nوضعیت شما"))
                                {
                                    counter++;
                                    Console.WriteLine("Doctor Change Status");
                                    data = DoctorChangeStatus(data, (string)xlRange.Cells[i, 2].Value2.ToString(),
                                        (string)xlRange.Cells[i, 4].Value2.ToString(), (string)xlRange.Cells[i, 5].Value2.ToString());
                                }
                                else if (partialConmmand.Equals("سامانه رسا\nکد کاربری"))
                                {
                                    counter++;
                                    Console.WriteLine("Charge Notification");
                                    data = ChargeNotification(data, (string)xlRange.Cells[i, 2].Value2.ToString(),
                                        (string)xlRange.Cells[i, 4].Value2.ToString(), (string)xlRange.Cells[i, 5].Value2.ToString());
                                }
                                else if (partialConmmand.Equals("سامانه رسا\nشما قبلاً"))
                                {
                                    counter++;
                                    Console.WriteLine("Patient Registration Conflict");
                                    data = PatientRegistrationConflict(data, (string)xlRange.Cells[i, 2].Value2.ToString(),
                                       (string)xlRange.Cells[i, 4].Value2.ToString(), (string)xlRange.Cells[i, 5].Value2.ToString());

                                }
                                else if (partialConmmand.Equals("به رسا خوش آمدید\n کد"))
                                {
                                    counter++;
                                    Console.WriteLine("Verify");
                                    data = Verify(data, (string)xlRange.Cells[i, 2].Value2.ToString(),
                                       (string)xlRange.Cells[i, 4].Value2.ToString(), (string)xlRange.Cells[i, 5].Value2.ToString());
                                }
                                else if ((partialConmmand.Equals("سامانه رسا\nپزشک گرام") || partialConmmand.Equals("سامانه رسا\nهمکار گرا")))
                                {
                                    counter++;
                                    Console.WriteLine("Doctor Periodic Report");
                                    data = DoctorPeriodicReport(data, (string)xlRange.Cells[i, 2].Value2.ToString(),
                                       (string)xlRange.Cells[i, 4].Value2.ToString(), (string)xlRange.Cells[i, 5].Value2.ToString());
                                }
                                else
                                {
                                    Console.WriteLine("/////////////////////////////////////////////// New Message ///////////////////////////////////////////////////////////");
                                }
                            }
                        }
                    }
                }
                ConvertToJson(data, k, month);
                Console.WriteLine(counter);
            }
        }
        public static List<Data> PatientRegistred(List<Data> data, string text, string number, string date)
        {
            int userCodeIndex = text.IndexOf("کد کاربری");
            int passwordIndex = text.IndexOf("رمز عبور");
            int lastPasswordIndex = text.IndexOf("برای ارتباط با پزشک خود");


            if (lastPasswordIndex == -1)
            {
                lastPasswordIndex = text.IndexOf("در نظر داشته باشید");
            }

            int userCodeLength = passwordIndex - userCodeIndex - 12;
            int passwordLength = lastPasswordIndex - passwordIndex - 12;

            string userCode = text.Substring(userCodeIndex + 11, userCodeLength);
            string password = text.Substring(passwordIndex + 10, passwordLength);

            Console.WriteLine(userCode);
            Console.WriteLine(password);
            data.Add(new Data
            {
                Type = "Patient Registred",
                Number = number,
                Date = date,
                UserCode = userCode,
                Password = password,
                MobileNumber = "",
                VerificationCode = "",
                DoctorCode = "",
                DoctorName = "",
                Charge = "",
                CallDurationHour = "",
                CallDurationMinute = "",
                CallDurationSecond = "",
                Cost = "",
                Credit = "",
                PeriodLength = "",
                NumberOfSuccessfulCall = "",
                NumberOfFailedCall = "",
                ConversationLength = "",
                Status = ""
            });
            
            return data;
        }
        public static List<Data> CallBockWithoutName(List<Data> data, string text, string number, string date)
        {
            int doctorCodeIndex = text.IndexOf("کد رسای پزشک شما");
            int lastDoctorCodeIndex = text.LastIndexOf("است");

            int doctorCodeLength = lastDoctorCodeIndex - doctorCodeIndex - 18;

            string doctorCode = text.Substring(doctorCodeIndex + 17, doctorCodeLength);

            Console.WriteLine(doctorCode);
            data.Add(new Data
            {
                Type = "Call Blocked Without Docter Name",
                Number = number,
                Date = date,
                UserCode = "",
                Password = "",
                MobileNumber = "",
                VerificationCode = "",
                DoctorCode = doctorCode,
                DoctorName = "",
                Charge = "",
                CallDurationHour = "",
                CallDurationMinute = "",
                CallDurationSecond = "",
                Cost = "",
                Credit = "",
                PeriodLength = "",
                NumberOfSuccessfulCall = "",
                NumberOfFailedCall = "",
                ConversationLength = "",
                Status = ""
            });
            return data;
        }
        public static List<Data> FeedbackAfterPromotionalCall(List<Data> data, string text, string number, string date)
        {
            int userCodeIndex = text.IndexOf("کد کاربری شما");
           
            int userCodeLength = text.Length - userCodeIndex - 16;

            string userCode = text.Substring(userCodeIndex + 16, userCodeLength);

            Console.WriteLine(userCode);
            data.Add(new Data
            {
                Type = "Feedback After Promotional Call",
                Number = number,
                Date = date,
                UserCode = userCode,
                Password = "",
                MobileNumber = "",
                VerificationCode = "",
                DoctorCode = "",
                DoctorName = "",
                Charge = "",
                CallDurationHour = "",
                CallDurationMinute = "",
                CallDurationSecond = "",
                Cost = "",
                Credit = "",
                PeriodLength = "",
                NumberOfSuccessfulCall = "",
                NumberOfFailedCall = "",
                ConversationLength = "",
                Status = ""
            });
            return data;
        }
        public static List<Data> Feedback(List<Data> data, string text, string number, string date)
        {
            int callDurationIndex = text.IndexOf("مدت تماس");
            int costIndex = text.IndexOf("هزینه‌");
            int creditIndex = text.IndexOf("اعتبار باقی مانده");
            int hourIndex = text.IndexOf("ساعت");
            int minuteIndex = text.IndexOf("دقیقه");
            int secondIndex = text.IndexOf("ثانیه");
            int unitIndex = text.IndexOf("تومان");
            int secondUnitIndex = text.LastIndexOf("تومان");

            string callDurationHour;
            string callDurationMinute;
            string callDurationSecond;

            if (hourIndex != -1)
            {
                callDurationHour = text.Substring(callDurationIndex + 10, hourIndex - callDurationIndex - 11);
                if (minuteIndex != -1)
                {
                    callDurationMinute = text.Substring(hourIndex + 7, minuteIndex - hourIndex - 8);
                    if (secondIndex != -1)
                        callDurationSecond = text.Substring(minuteIndex + 8, secondIndex - minuteIndex - 9);
                    else
                        callDurationSecond = "0";
                }
                else
                {
                    callDurationMinute = "0";
                    if (secondIndex != -1)
                        callDurationSecond = text.Substring(hourIndex + 7, secondIndex - hourIndex - 8);
                    else
                        callDurationSecond = "0";
                }
            }
            else
            {
                callDurationHour = "0";
                if (minuteIndex != -1)
                {
                    callDurationMinute = text.Substring(callDurationIndex + 10, minuteIndex - callDurationIndex - 11);
                    if (secondIndex != -1)
                        callDurationSecond = text.Substring(minuteIndex + 8, secondIndex - minuteIndex - 9);
                    else
                        callDurationSecond = "0";
                }
                else
                {
                    callDurationMinute = "0";
                    if (secondIndex != -1)
                        callDurationSecond = text.Substring(callDurationIndex + 10, secondIndex - callDurationIndex - 11);
                    else
                        callDurationSecond = "0";
                }   
            }

            string cost = text.Substring(costIndex + 8, unitIndex - costIndex - 9);
            string credit;
            if (text.Substring(creditIndex + 20, 3).Equals("صفر"))
                credit = "0";
            else
                credit = text.Substring(creditIndex + 20, secondUnitIndex - creditIndex - 22);
            if (text[secondUnitIndex - 2] == '-')
            {
                credit = "-" + credit;
            }
                
            Console.WriteLine(callDurationHour);
            Console.WriteLine(callDurationMinute);
            Console.WriteLine(callDurationSecond);
            Console.WriteLine(cost);
            Console.WriteLine(credit);
            data.Add(new Data
            {
                Type = "Feedback",
                Number = number,
                Date = date,
                UserCode = "",
                Password = "",
                MobileNumber = "",
                VerificationCode = "",
                DoctorCode = "",
                DoctorName = "",
                Charge = "",
                CallDurationHour = callDurationHour,
                CallDurationMinute = callDurationMinute,
                CallDurationSecond = callDurationSecond,
                Cost = cost,
                Credit = credit,
                PeriodLength = "",
                NumberOfSuccessfulCall = "",
                NumberOfFailedCall = "",
                ConversationLength = "",
                Status = ""
            });
            return data;
        }

        public static List<Data> CallBlockWithName(List<Data> data, string text, string number, string date)
        {
            string doctorCode;
            string doctorName;
            if (text.IndexOf(")") == -1)
            {
                int doctorCodeIndex = text.IndexOf("کد رسای پزشک شما");
                int doctorNameIndex = text.IndexOf("دکتر");
                int doctorNameOffset = 5;
                int doctorNameSizeOffset = 6;
                if (doctorNameIndex == -1)
                {
                    doctorNameIndex = text.IndexOf("پزشک شما");
                    doctorNameOffset = 10;
                    doctorNameSizeOffset = 11;

                }
                int lastDoctorNameIndex = text.IndexOf("تنها");
                int lastDoctorCodeIndex = text.LastIndexOf("است");

                int doctorCodeLength = lastDoctorCodeIndex - doctorCodeIndex - 18;
                int doctorNameLength = lastDoctorNameIndex - doctorNameIndex - doctorNameSizeOffset;

                doctorCode = text.Substring(doctorCodeIndex + 17, doctorCodeLength);
                doctorName = text.Substring(doctorNameIndex + doctorNameOffset, doctorNameLength);
            }
            else
            {
                if(text.IndexOf("ایشان") != -1)
                    return CallBlocked(data, text, number, date);
                else
                    return CallBlockedFirstTime(data, text, number, date);

            }
            Console.WriteLine(doctorCode);
            Console.WriteLine(doctorName);
            data.Add(new Data
            {
                Type = "Call Blocked With Docter Name",
                Number = number,
                Date = date,
                UserCode = "",
                Password = "",
                MobileNumber = "",
                VerificationCode = "",
                DoctorCode = doctorCode,
                DoctorName = doctorName,
                Charge = "",
                CallDurationHour = "",
                CallDurationMinute = "",
                CallDurationSecond = "",
                Cost = "",
                Credit = "",
                PeriodLength = "",
                NumberOfSuccessfulCall = "",
                NumberOfFailedCall = "",
                ConversationLength = "",
                Status = ""
            });
            return data;
        }
        public static List<Data> OneTimeChargeNotification(List<Data> data, string text, string number, string date)
        {
            int chargeIndex = text.IndexOf("مبلغ");
            int unitIndeex = text.IndexOf("تومان");

            int chargeLength = unitIndeex - chargeIndex - 6;

            string charge = text.Substring(chargeIndex + 5, chargeLength);

            Console.WriteLine(charge);
            data.Add(new Data
            {
                Type = "One Time Charge Notification",
                Number = number,
                Date = date,
                UserCode = "",
                Password = "",
                MobileNumber = "",
                VerificationCode = "",
                DoctorCode = "",
                DoctorName = "",
                Charge = charge,
                CallDurationHour = "",
                CallDurationMinute = "",
                CallDurationSecond = "",
                Cost = "",
                Credit = "",
                PeriodLength = "",
                NumberOfSuccessfulCall = "",
                NumberOfFailedCall = "",
                ConversationLength = "",
                Status = ""
            });
            return data;
        }
        public static List<Data> DoctorChangeStatus(List<Data> data , string text, string number, string date)
        {
            int startStatusIndex = text.IndexOf("به");
            int finishStatusIndex = text.IndexOf("تغییر");

            int chargeLength = finishStatusIndex - startStatusIndex - 4;

            string status = text.Substring(startStatusIndex + 3, chargeLength);

            Console.WriteLine(status);
            data.Add(new Data
            {
                Type = "Doctor Change Status",
                Number = number,
                Date = date,
                UserCode = "",
                Password = "",
                MobileNumber = "",
                VerificationCode = "",
                DoctorCode = "",
                DoctorName = "",
                Charge = "",
                CallDurationHour = "",
                CallDurationMinute = "",
                CallDurationSecond = "",
                Cost = "",
                Credit = "",
                PeriodLength = "",
                NumberOfSuccessfulCall = "",
                NumberOfFailedCall = "",
                ConversationLength = "",
                Status = status
            });
            return data;
        }
        public static List<Data> ChargeNotification(List<Data> data, string text, string number, string date)
        {
            int userCodeIndex = text.IndexOf("کاربری");
            int chargeIndex = text.IndexOf("به مبلغ");
            int firstUnitIndex = text.IndexOf("تومان");
            int creditIndex = text.IndexOf("موجودی کنونی");
            int lastUnitIndex = text.LastIndexOf("تومان");

            int userCodeLength = chargeIndex - userCodeIndex - 8;
            int chargeLength = firstUnitIndex - chargeIndex - 9;
            int creditLength = lastUnitIndex - creditIndex - 15;

            string userCode = text.Substring(userCodeIndex + 7, userCodeLength);
            string charge = text.Substring(chargeIndex + 8, chargeLength);
            string credit = text.Substring(creditIndex + 14, creditLength);

            Console.WriteLine(userCode);
            Console.WriteLine(charge);
            Console.WriteLine(credit);
            data.Add(new Data
            {
                Type = "Charge Notification",
                Number = number,
                Date = date,
                UserCode = userCode,
                Password = "",
                MobileNumber = "",
                VerificationCode = "",
                DoctorCode = "",
                DoctorName = "",
                Charge = charge,
                CallDurationHour = "",
                CallDurationMinute = "",
                CallDurationSecond = "",
                Cost = "",
                Credit = credit,
                PeriodLength = "",
                NumberOfSuccessfulCall = "",
                NumberOfFailedCall = "",
                ConversationLength = "",
                Status = ""
            });
            return data;
        }
        public static List<Data> PatientRegistrationConflict(List<Data> data, string text, string number, string date)
        {
            int mobileNumbeIndex = text.IndexOf("موبایل");
            int lastMobileNumberIndex = text.IndexOf("در سامانه");
            int userCodeIndex = text.IndexOf("کد کاربری");
            int passwordIndex = text.IndexOf("رمز عبور");

            int mobileNumberLength = lastMobileNumberIndex - mobileNumbeIndex - 10;
            int userCodeLength = passwordIndex - userCodeIndex - 12;
            int passwordLength = text.Length - passwordIndex - 10;
           
            string mobileNumber = text.Substring(mobileNumbeIndex + 8, mobileNumberLength);
            string userCode = text.Substring(userCodeIndex + 11, userCodeLength);
            string password = text.Substring(passwordIndex + 10, passwordLength);

            Console.WriteLine(mobileNumber);
            Console.WriteLine(userCode);
            Console.WriteLine(password);
            data.Add(new Data
            {
                Type = "Patient Registration Conflict",
                Number = number,
                Date = date,
                UserCode = userCode,
                Password = password,
                MobileNumber = mobileNumber,
                VerificationCode = "",
                DoctorCode = "",
                DoctorName = "",
                Charge = "",
                CallDurationHour = "",
                CallDurationMinute = "",
                CallDurationSecond = "",
                Cost = "",
                Credit = "",
                PeriodLength = "",
                NumberOfSuccessfulCall = "",
                NumberOfFailedCall = "",
                ConversationLength = "",
                Status = ""
            });
            return data;
        }
        public static List<Data> Verify(List<Data> data, string text, string number, string date)
        {
            int verificationCodeIndex = text.IndexOf("همراه");

            int verificationCodeLength = text.Length - verificationCodeIndex - 9;

            string verificationCode = text.Substring(verificationCodeIndex + 8, verificationCodeLength);

            Console.WriteLine(verificationCode);
            data.Add(new Data
            {
                Type = "Verify",
                Number = number,
                Date = date,
                UserCode = "",
                Password = "",
                MobileNumber = "",
                VerificationCode = verificationCode,
                DoctorCode = "",
                DoctorName = "",
                Charge = "",
                CallDurationHour = "",
                CallDurationMinute = "",
                CallDurationSecond = "",
                Cost = "",
                Credit = "",
                PeriodLength = "",
                NumberOfSuccessfulCall = "",
                NumberOfFailedCall = "",
                ConversationLength = "",
                Status = ""
            });
            return data;
        }
        public static List<Data> DoctorPeriodicReport(List<Data> data, string text, string number, string date)
        {
            int periodLengthSizeOffset = 19;
            int periodLengthOffset = 18;
            int conversationLengthOffset = 8;
            int periodLengthIndex = text.IndexOf("پزشک گرامی");
            if (periodLengthIndex == -1)
            {
                periodLengthSizeOffset = 20;
                periodLengthOffset = 19;
                periodLengthIndex = text.IndexOf("همکار گرامی");
            }
               
            int lastPeriodLengthIndex = text.IndexOf("گذشته");
            int successfulCallIndex = text.IndexOf("تماس موفق");
            int failedCallIndex = text.IndexOf("تماس ناموفق");
            int conversationLengthIndex = text.IndexOf("بیماران");
            int unitIndex = text.IndexOf("دقیقه");

            int periodLengthSize = lastPeriodLengthIndex - periodLengthIndex - periodLengthSizeOffset;
            int successfulCallLength = successfulCallIndex - lastPeriodLengthIndex - 8; 
            int failedCallLength = failedCallIndex - successfulCallIndex - 13;
            int ConversationLengthSize = unitIndex - conversationLengthIndex - 9;

            if (text.IndexOf("نیز") != -1)
            {
                ConversationLengthSize -= 4;
                conversationLengthOffset += 4;
            }

            string periodLength = text.Substring(periodLengthIndex + periodLengthOffset, periodLengthSize);
            string numberOfSuccessfulCall = text.Substring(lastPeriodLengthIndex + 6, successfulCallLength);
            string numberOfFailedCall = text.Substring(successfulCallIndex + 12, failedCallLength);
            string conversationLenth = text.Substring(conversationLengthIndex + conversationLengthOffset, ConversationLengthSize);
            
            Console.WriteLine(periodLength);
            Console.WriteLine(numberOfSuccessfulCall);
            Console.WriteLine(numberOfFailedCall);
            Console.WriteLine(conversationLenth);
            data.Add(new Data
            {
                Type = "Doctor Periodic Report",
                Number = number,
                Date = date,
                UserCode = "",
                Password = "",
                MobileNumber = "",
                VerificationCode = "",
                DoctorCode = "",
                DoctorName = "",
                Charge = "",
                CallDurationHour = "",
                CallDurationMinute = "",
                CallDurationSecond = "",
                Cost = "",
                Credit = "",
                PeriodLength = periodLength,
                NumberOfSuccessfulCall = numberOfSuccessfulCall,
                NumberOfFailedCall = numberOfFailedCall,
                ConversationLength = conversationLenth,
                Status = ""
            });
            return data;
        }
        public static List<Data> CallBlockedFirstTime(List<Data> data, string text, string number, string date)
        {
            int doctorNameSizeOffset = 7;
            int doctorNameOffset = 5;
            int doctorNameIndex = text.IndexOf("دکتر");
            if (doctorNameIndex == -1)
            {
                doctorNameIndex = text.IndexOf("تلفنی");
                doctorNameSizeOffset = 11;
                doctorNameOffset = 9;
            }
            int doctorCodeIndex = text.IndexOf("کد");
            int lastDoctorCodeIndex = text.IndexOf("تنها");

            int doctorCodeLength = lastDoctorCodeIndex - doctorCodeIndex - 5;
            int doctorNameLength = doctorCodeIndex - doctorNameIndex - doctorNameSizeOffset;

            string doctorCode = text.Substring(doctorCodeIndex + 3, doctorCodeLength);
            string doctorName = text.Substring(doctorNameIndex + doctorNameOffset, doctorNameLength);

            Console.WriteLine(doctorCode);
            Console.WriteLine(doctorName);
            data.Add(new Data
            {
                Type = "Call Blocked First Time",
                Number = number,
                Date = date,
                UserCode = "",
                Password = "",
                MobileNumber = "",
                VerificationCode = "",
                DoctorCode = doctorCode,
                DoctorName = doctorName,
                Charge = "",
                CallDurationHour = "",
                CallDurationMinute = "",
                CallDurationSecond = "",
                Cost = "",
                Credit = "",
                PeriodLength = "",
                NumberOfSuccessfulCall = "",
                NumberOfFailedCall = "",
                ConversationLength = "",
                Status = ""
            });
            return data;
        }
        public static List<Data> CallBlocked(List<Data> data, string text, string number, string date)
        {
            int doctorNameSizeOffset = 6;
            int doctorNameOffset = 5;
            int doctorNameIndex = text.IndexOf("دکتر");
            if (doctorNameIndex == -1)
            {
                doctorNameIndex = text.IndexOf("تلفنی");
                doctorNameSizeOffset = 10;
                doctorNameOffset = 9;
            }
            int doctorCodeIndex = text.IndexOf("کد ایشان");
            int lastDoctorCodeIndex = text.IndexOf("را وارد");
            int lastDoctorNameIndex = text.IndexOf("تنها");

            int doctorCodeLength = lastDoctorCodeIndex - doctorCodeIndex - 12;
            int doctorNameLength = lastDoctorNameIndex - doctorNameIndex - doctorNameSizeOffset;

            string doctorCode = text.Substring(doctorCodeIndex + 10, doctorCodeLength);
            string doctorName = text.Substring(doctorNameIndex + doctorNameOffset, doctorNameLength);

            Console.WriteLine(doctorCode);
            Console.WriteLine(doctorName);
            data.Add(new Data
            {
                Type = "Call Blocked",
                Number = number,
                Date = date,
                UserCode = "",
                Password = "",
                MobileNumber = "",
                VerificationCode = "",
                DoctorCode = doctorCode,
                DoctorName = doctorName,
                Charge = "",
                CallDurationHour = "",
                CallDurationMinute = "",
                CallDurationSecond = "",
                Cost = "",
                Credit = "",
                PeriodLength = "",
                NumberOfSuccessfulCall = "",
                NumberOfFailedCall = "",
                ConversationLength = "",
                Status = ""
            });
            return data;
        }
    }
}
