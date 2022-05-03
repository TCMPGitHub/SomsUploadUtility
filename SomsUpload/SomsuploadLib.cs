using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualBasic;

namespace SomsUploadLib
{
    public class SomsUploadLib : IComponent
    {
        public ISite Site
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public string DBBassConn { get; set;}
        public string DBPatsConn { get; set; }

        //= "Data Source = APO7LTCMP006512; Initial Catalog = BassWebTest; Integrated Security = True; MultipleActiveResultSets=True;Application Name = EntityFramework";

        public string FilePath{ get; set; }
        public string ErrorMessage { get; set; }
        public bool Filtered { get; set; }
        //public object ConfigurationManager { get; private set; }

        public event EventHandler Disposed;

        //public int Import() {
        //    GetAmountOfCopiedRecords(216);
        //    return 216;
        //}

        //public int Import()
        //{
        //    int nextId = GeMaxUploadId(Path.GetFileName(FilePath));
        //    LogWriter.LogMessageToFile("SomsUploadID is "+ nextId.ToString());

        //    string conditionExpression = "SCHEDULEDRELEASEDATE IS NOT null And SCHEDULEDRELEASEDATE <> '' And LOC Not like 'SACCO%' And LOC Not like 'RENT%' And LOC Not like 'LPU%' And LOC Not like 'COCF%'";
        //    string CSVFileConnectionString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"text;HDR=Yes;FMT=Delimited\";", Path.GetDirectoryName(FilePath));

        //    try
        //    {
        //        DataTable dt = new DataTable();
        //        using (OleDbConnection con = new OleDbConnection(CSVFileConnectionString))
        //        {
        //            con.Open();
        //            var csvQuery = string.Format("select * from [{0}]", Path.GetFileName(FilePath));
        //            using (OleDbDataAdapter da = new OleDbDataAdapter(csvQuery, con))
        //            {
        //                da.Fill(dt);
        //            }

        //            //need to remove all row that has ScheduleReleaseDate is null or blank, LOC is not start with SACCO, RENT, LUP 
        //            if (Filtered)
        //            {
        //                DataRow[] dw = dt.Select(conditionExpression);
        //                dt = dw.CopyToDataTable();
        //            }

        //            dt.Columns.Add("SomsUploadID", typeof(System.Int32));
        //            foreach (DataRow row in dt.Rows)
        //            {
        //                row["SomsUploadID"] = nextId;   // set it to nextid
        //            }
        //        }

        //        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(DBBassConn))
        //        {
        //            bulkCopy.ColumnMappings.Add("SomsUploadID", "SomsUploadID");
        //            bulkCopy.ColumnMappings.Add("OFFENDERID", "SomsOffenderID");
        //            bulkCopy.ColumnMappings.Add("UNIT", "Unit");
        //            bulkCopy.ColumnMappings.Add("LOC", "Loc");
        //            bulkCopy.ColumnMappings.Add("SCHEDULEDRELEASEDATE", "ScheduledReleaseDate");
        //            bulkCopy.ColumnMappings.Add("HOLD", "Hold");
        //            bulkCopy.ColumnMappings.Add("DETAINERAUTHTYPE", "DetainerAuthType");
        //            bulkCopy.ColumnMappings.Add("CDCNUMBER", "CDCNumber");
        //            bulkCopy.ColumnMappings.Add("INMATELASTNAME", "InmateLastName");
        //            bulkCopy.ColumnMappings.Add("INMATESEXCODE", "InmateSexCode");
        //            bulkCopy.ColumnMappings.Add("INMATEFIRSTNAME", "InmateFirstName");
        //            bulkCopy.ColumnMappings.Add("INMATESTATUSCODE", "InmateStatusCode");
        //            bulkCopy.ColumnMappings.Add("PRIMARYBEDUSE", "PrimaryBedUse");
        //            bulkCopy.ColumnMappings.Add("CELLBED", "CellBed");
        //            bulkCopy.ColumnMappings.Add("CURRENTINCARCBEGINDATE", "CurrentIncarcBeginDate");
        //            bulkCopy.ColumnMappings.Add("MHDESC", "MhDesc");
        //            bulkCopy.ColumnMappings.Add("SSNUM", "SSNum");
        //            bulkCopy.ColumnMappings.Add("COUNTYOFSENTENCING", "CountyOfSentencing");
        //            bulkCopy.ColumnMappings.Add("COMMITMENTTYPE", "CommitmentType");
        //            bulkCopy.ColumnMappings.Add("FILEDATE", "FileDate");
        //            bulkCopy.ColumnMappings.Add("INMATEDATEOFBIRTH", "InmateDateOfBirth");
        //            bulkCopy.ColumnMappings.Add("DEVELOPDISABLEDEVALUATION", "DevelopDisabledEvaluation");
        //            bulkCopy.ColumnMappings.Add("DPPHEARINGCODE", "DPPHearingCode");
        //            bulkCopy.ColumnMappings.Add("DPPMOBILITYCODE", "DPPMobilityCode");
        //            bulkCopy.ColumnMappings.Add("DPPSPEECHCODE", "DPPSpeechCode");
        //            bulkCopy.ColumnMappings.Add("DPPVISIONCODE", "DPPVisionCode");
        //            bulkCopy.ColumnMappings.Add("DAYSLEFT", "DaysLeft");
        //            bulkCopy.ColumnMappings.Add("LIFER", "Lifer");
        //            bulkCopy.ColumnMappings.Add("COUNTYOFRELEASE", "CountyOfRelease");
        //            bulkCopy.ColumnMappings.Add("MENTALHEALTHTREATMENTNEEDED", "MentalHealthTreatmentNeeded");
        //            bulkCopy.ColumnMappings.Add("PAROLEUNIT", "PAROLEUNIT");
        //            bulkCopy.ColumnMappings.Add("PAROLEAGENTLASTNAME__", "PAROLEAGENTLASTNAME");
        //            bulkCopy.ColumnMappings.Add("PAROLEAGENTFIRSTNAME", "PAROLEAGENTFIRSTNAME");
        //            bulkCopy.ColumnMappings.Add("PAROLEAGENTMIDDLENAME", "PAROLEAGENTMIDDLENAME");

        //            bulkCopy.DestinationTableName = "SomsRecord";
        //            bulkCopy.BatchSize = 0;
        //            bulkCopy.WriteToServer(dt);
        //            bulkCopy.Close();
        //        }

        //        LogWriter.LogMessageToFile("SomsUpload bulkCopy Complated");
        //        GetAmountOfCopiedRecords(nextId);
        //        return nextId;
        //    }
        //    catch (Exception ex)
        //    {
        //        LogWriter.LogMessageToFile(ex.Message);
        //        return -1;
        //    }
        //}
        public int Import()
        {
            int nextId = GeMaxUploadId(Path.GetFileName(FilePath));
            LogWriter.LogMessageToFile("SomsUploadID is " + nextId.ToString());

            string conditionExpression = "SCHEDULEDRELEASEDATE IS NOT null And SCHEDULEDRELEASEDATE <> '' And LOC Not like 'SACCO%' And LOC Not like 'RENT%' And LOC Not like 'LPU%' And LOC Not like 'COCF%'";
            string CSVFileConnectionString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"text;HDR=Yes;FMT=Delimited\";", Path.GetDirectoryName(FilePath));

            try
            {
                DataTable dt = new DataTable();
                using (OleDbConnection con = new OleDbConnection(CSVFileConnectionString))
                {
                    con.Open();
                    var csvQuery = string.Format("select * from [{0}]", Path.GetFileName(FilePath));
                    using (OleDbDataAdapter da = new OleDbDataAdapter(csvQuery, con))
                    {
                        da.Fill(dt);
                    }

                    //need to remove all row that has ScheduleReleaseDate is null or blank, LOC is not start with SACCO, RENT, LUP 
                    if (Filtered)
                    {
                        DataRow[] dw = dt.Select(conditionExpression);
                        dt = dw.CopyToDataTable();
                    }

                    dt.Columns.Add("SomsUploadID", typeof(System.Int32));
                    foreach (DataRow row in dt.Rows)
                    {
                        row["SomsUploadID"] = nextId;   // set it to nextid
                    }
                }

                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(DBBassConn))
                {
                    bulkCopy.ColumnMappings.Add("SomsUploadID", "SomsUploadID");
                    bulkCopy.ColumnMappings.Add("OFFENDERID", "SomsOffenderID");
                    bulkCopy.ColumnMappings.Add("UNIT", "Unit");
                    bulkCopy.ColumnMappings.Add("LOC", "Loc");
                    bulkCopy.ColumnMappings.Add("SCHEDULEDRELEASEDATE", "ScheduledReleaseDate");
                    bulkCopy.ColumnMappings.Add("HOLD", "Hold");
                    bulkCopy.ColumnMappings.Add("DETAINERAUTHTYPE", "DetainerAuthType");
                    bulkCopy.ColumnMappings.Add("CDCNUMBER", "CDCNumber");
                    bulkCopy.ColumnMappings.Add("INMATELASTNAME", "InmateLastName");
                    bulkCopy.ColumnMappings.Add("INMATESEXCODE", "InmateSexCode");
                    bulkCopy.ColumnMappings.Add("INMATEFIRSTNAME", "InmateFirstName");
                    bulkCopy.ColumnMappings.Add("INMATESTATUSCODE", "InmateStatusCode");
                    bulkCopy.ColumnMappings.Add("PRIMARYBEDUSE", "PrimaryBedUse");
                    bulkCopy.ColumnMappings.Add("CELLBED", "CellBed");
                    bulkCopy.ColumnMappings.Add("CURRENTINCARCBEGINDATE", "CurrentIncarcBeginDate");
                    bulkCopy.ColumnMappings.Add("MHDESC", "MhDesc");
                    bulkCopy.ColumnMappings.Add("SSNUM", "SSNum");
                    bulkCopy.ColumnMappings.Add("COUNTYOFSENTENCING", "CountyOfSentencing");
                    bulkCopy.ColumnMappings.Add("COMMITMENTTYPE", "CommitmentType");
                    bulkCopy.ColumnMappings.Add("FILEDATE", "FileDate");
                    bulkCopy.ColumnMappings.Add("INMATEDATEOFBIRTH", "InmateDateOfBirth");
                    bulkCopy.ColumnMappings.Add("DEVELOPDISABLEDEVALUATION", "DevelopDisabledEvaluation");
                    bulkCopy.ColumnMappings.Add("DPPHEARINGCODE", "DPPHearingCode");
                    bulkCopy.ColumnMappings.Add("DPPMOBILITYCODE", "DPPMobilityCode");
                    bulkCopy.ColumnMappings.Add("DPPSPEECHCODE", "DPPSpeechCode");
                    bulkCopy.ColumnMappings.Add("DPPVISIONCODE", "DPPVisionCode");
                    bulkCopy.ColumnMappings.Add("DAYSLEFT", "DaysLeft");
                    bulkCopy.ColumnMappings.Add("LIFER", "Lifer");
                    bulkCopy.ColumnMappings.Add("COUNTYOFRELEASE", "CountyOfRelease");
                    bulkCopy.ColumnMappings.Add("MENTALHEALTHTREATMENTNEEDED", "MentalHealthTreatmentNeeded");
                    bulkCopy.ColumnMappings.Add("PAROLEUNIT", "PAROLEUNIT");
                    bulkCopy.ColumnMappings.Add("PAROLEAGENTLASTNAME__", "PAROLEAGENTLASTNAME");
                    bulkCopy.ColumnMappings.Add("PAROLEAGENTFIRSTNAME", "PAROLEAGENTFIRSTNAME");
                    bulkCopy.ColumnMappings.Add("PAROLEAGENTMIDDLENAME", "PAROLEAGENTMIDDLENAME");

                    bulkCopy.DestinationTableName = "SomsRecord";
                    bulkCopy.BatchSize = 0;
                    bulkCopy.BulkCopyTimeout = 2000;
                    bulkCopy.WriteToServer(dt);
                    bulkCopy.Close();
                }

                LogWriter.LogMessageToFile("SomsUpload bulkCopy to BASS Complated");

                using (SqlBulkCopy patsbulkCopy = new SqlBulkCopy(DBPatsConn))
                {
                    patsbulkCopy.DestinationTableName = "SomsRecord";
                    patsbulkCopy.ColumnMappings.Add("SomsUploadID", "SomsUploadID");
                    patsbulkCopy.ColumnMappings.Add("OFFENDERID", "SomsOffenderID");
                    patsbulkCopy.ColumnMappings.Add("UNIT", "Unit");
                    patsbulkCopy.ColumnMappings.Add("LOC", "Loc");
                    patsbulkCopy.ColumnMappings.Add("SCHEDULEDRELEASEDATE", "ScheduledReleaseDate");
                    patsbulkCopy.ColumnMappings.Add("CDCNUMBER", "CDCNumber");
                    patsbulkCopy.ColumnMappings.Add("INMATELASTNAME", "InmateLastName");
                    patsbulkCopy.ColumnMappings.Add("INMATESEXCODE", "InmateSexCode");
                    patsbulkCopy.ColumnMappings.Add("INMATEFIRSTNAME", "InmateFirstName");
                    patsbulkCopy.ColumnMappings.Add("INMATESTATUSCODE", "InmateStatusCode");
                    patsbulkCopy.ColumnMappings.Add("MHDESC", "MhDesc");
                    patsbulkCopy.ColumnMappings.Add("SSNUM", "SSNum");
                    patsbulkCopy.ColumnMappings.Add("COUNTYOFSENTENCING", "CountyOfSentencing");
                    patsbulkCopy.ColumnMappings.Add("COMMITMENTTYPE", "CommitmentType");
                    patsbulkCopy.ColumnMappings.Add("INMATEDATEOFBIRTH", "InmateDateOfBirth");
                    patsbulkCopy.ColumnMappings.Add("DEVELOPDISABLEDEVALUATION", "DevelopDisabledEvaluation");
                    patsbulkCopy.ColumnMappings.Add("DPPHEARINGCODE", "DPPHearingCode");
                    patsbulkCopy.ColumnMappings.Add("DPPMOBILITYCODE", "DPPMobilityCode");
                    patsbulkCopy.ColumnMappings.Add("DPPSPEECHCODE", "DPPSpeechCode");
                    patsbulkCopy.ColumnMappings.Add("DPPVISIONCODE", "DPPVisionCode");
                    patsbulkCopy.ColumnMappings.Add("DAYSLEFT", "DaysLeft");
                    patsbulkCopy.ColumnMappings.Add("LIFER", "Lifer");
                    patsbulkCopy.ColumnMappings.Add("COUNTYOFRELEASE", "CountyOfRelease");
                    patsbulkCopy.ColumnMappings.Add("MENTALHEALTHTREATMENTNEEDED", "MentalHealthTreatmentNeeded");
                    patsbulkCopy.ColumnMappings.Add("PAROLEUNIT", "PAROLEUNIT");
                    patsbulkCopy.ColumnMappings.Add("PAROLEAGENTLASTNAME__", "PAROLEAGENTLASTNAME");
                    patsbulkCopy.ColumnMappings.Add("PAROLEAGENTFIRSTNAME", "PAROLEAGENTFIRSTNAME");
                    patsbulkCopy.ColumnMappings.Add("PAROLEAGENTMIDDLENAME", "PAROLEAGENTMIDDLENAME");
                    patsbulkCopy.ColumnMappings.Add("FILEDATE", "FileDate");
                    patsbulkCopy.DestinationTableName = "SomsRecord";
                    patsbulkCopy.BatchSize = 0;
                    patsbulkCopy.BulkCopyTimeout = 2000;
                    patsbulkCopy.WriteToServer(dt);
                    patsbulkCopy.Close();
                }
                LogWriter.LogMessageToFile("SomsUpload bulkCopy to PATS Complated");

                GetAmountOfCopiedRecords(nextId);
                return nextId;
            }
            catch (Exception ex)
            {
                LogWriter.LogMessageToFile(ex.Message);
                return -1;
            }
        }

        private int GetAmountOfCopiedRecords(int nextId)
        {
            try
            {
                using (SqlConnection cnz = new SqlConnection(DBBassConn))
                {
                    SqlCommand cmdzs = new SqlCommand("Select Count(*) From SomsRecord where SomsUploadId = " + nextId, cnz);
                    cmdzs.CommandType = CommandType.Text;
                    cnz.Open();
                    var obj = cmdzs.ExecuteScalar();
                    LogWriter.LogMessageToFile("SomsUploadId " + nextId + " total records " + obj + " found");
                    return obj != null ? int.Parse(obj.ToString()) : 0;
                }
            }
            catch (Exception ex)
            {
                LogWriter.LogMessageToFile(ex.Message);
                return -1;
            }
        }

        public string ProcessDataMappingBass(int nextId)
        {
            string msg = "Import For Bass...";
            LogWriter.LogMessageToFile(msg);
            
            //call sp to precess the matches
            using (SqlConnection con = new SqlConnection(DBBassConn))
            {
                //Create the SqlCommand object
                SqlCommand cmd = new SqlCommand("spImportSomsRecord", con);

                //Specify that the SqlCommand is a stored procedure
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 2000;
                //Add the input parameters to the command object
                cmd.Parameters.AddWithValue("@SomsOffenderId", 0);
                cmd.Parameters.AddWithValue("@SomsUploadId", nextId);

                //Add the output parameter to the command object
                SqlParameter outPutParameter = new SqlParameter();
                outPutParameter.ParameterName = "@ResultMessage";
                outPutParameter.SqlDbType = System.Data.SqlDbType.NVarChar;
                outPutParameter.Size = 8000;
                outPutParameter.Direction = System.Data.ParameterDirection.Output;
                cmd.Parameters.Add(outPutParameter);

                //Open the connection and execute the query
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
                //Retrieve the value of the output parameter
                string message = outPutParameter.Value.ToString();
                LogWriter.LogMessageToFile(message);
                msg = msg + Environment.NewLine + message;
            }
            return msg;
        }
        public string ProcessDataMappingPats(int nextId)
        {
            string msg = "Import For Pats...";
            LogWriter.LogMessageToFile(msg);
            //do PATS
            using (SqlConnection con1 = new SqlConnection(DBPatsConn))
            {
                //Create the SqlCommand object
                SqlCommand cmd = new SqlCommand("spImportSomsRecordToDevPats", con1);

                //Specify that the SqlCommand is a stored procedure
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 2000;
                //Add the input parameters to the command object
                cmd.Parameters.AddWithValue("@SomsOffenderId", 0);
                cmd.Parameters.AddWithValue("@SomsUploadId", nextId);

                //Add the output parameter to the command object
                SqlParameter outPutParameter = new SqlParameter();
                outPutParameter.ParameterName = "@ResultMessage";
                outPutParameter.SqlDbType = System.Data.SqlDbType.NVarChar;
                outPutParameter.Size = 8000;
                outPutParameter.Direction = System.Data.ParameterDirection.Output;
                cmd.Parameters.Add(outPutParameter);

                //Open the connection and execute the query
                con1.Open();
                cmd.ExecuteNonQuery();
                con1.Close();
                //Retrieve the value of the output parameter
                string message = outPutParameter.Value.ToString();
                LogWriter.LogMessageToFile(message);
                msg = msg + Environment.NewLine + message;
            }
            return msg;
        }

        private int GeMaxUploadId(string FileName)
        {
            try {
                using (SqlConnection cnz = new SqlConnection(DBBassConn))
                {
                    using (SqlCommand cmd = new SqlCommand("INSERT INTO SomsUpload(OrigFileName, ContentType, CreatedDate) VALUES(@file,@type, @date); SELECT SCOPE_IDENTITY();", cnz))
                    {
                        cmd.Parameters.AddWithValue("@file", FileName);
                        cmd.Parameters.AddWithValue("@type", "application/vnd.ms-excel");
                        cmd.Parameters.AddWithValue("@date", DateTime.Now);
                        cnz.Open();

                        int newID = Convert.ToInt32(cmd.ExecuteScalar());

                        // int newID = cmd.ExecuteNonQuery();
                        // var obj = cmd.ExecuteScalar();
                        if (cnz.State == System.Data.ConnectionState.Open)
                            cnz.Close();
                        LogWriter.LogMessageToFile("Insert to SomsUpload");
                        return newID;
                    
                    }
                }
            }
            catch (Exception ex)
            {
                LogWriter.LogMessageToFile(ex.Message);
                return -1;
            }        
        }

        // IComponent extends IDisposable.
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free other state (managed objects).
            }
            // Free your own state (unmanaged objects).
        }

        ~SomsUploadLib()
        {
            Dispose(false);
        }
    }
}
