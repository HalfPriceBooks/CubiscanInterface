using System;
using System.Linq;
using CubiscanInterface.DBHelpers;
using CubiscanInterface.Models;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using Dapper;
using OfficeOpenXml;
using System.IO;
using System.Configuration;
using System.Text;


namespace CubiscanInterface.DBHelpers
{
    class HPBInterface : IDisposable
    {
        #region "Dispose Logic"
        private bool disposed = false;
        public string stringConn;

        /// <summary>
        /// Constructor
        /// </summary>
        public HPBInterface()
        {
            CubiscanInterface.DBHelpers.DBConnection dbc = new DBHelpers.DBConnection();
            // Determine running mode 
            stringConn = dbc.ReadSetting((dbc.ReadSetting("MODE")).ToUpper().ToString());
            Dispose(false);
        }

        /// <summary>
        /// Implements Dipose method
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Overrides Dispose method
        /// </summary>
        protected virtual void Dispose(bool disposeManagedResources)
        {
            if (!this.disposed)
            {
                disposed = true;
            }
        }

        #endregion

        public int Upload_Dimensions(UploadParams _uploadParams)
        {
            int _res = 0;
            string xItem = "";
            try
            {
                DynamicParameters param = new DynamicParameters();
                xItem = _uploadParams.ITEM;
                param.Add("@ITEM", xItem);
                param.Add("@QUANTITY_UM", _uploadParams.QUANTITY_UM);
                param.Add("@COMPANY", _uploadParams.COMPANY);
                param.Add("@LENGTH", _uploadParams.LENGTH);
                param.Add("@WIDTH", _uploadParams.WIDTH);
                param.Add("@HEIGHT", _uploadParams.HEIGHT);
                param.Add("@DIMENSION_UM", _uploadParams.DIMENSION_UM);
                param.Add("@DIM_WEIGHT", _uploadParams.DIM_WEIGHT);
                param.Add("@VOLUME", _uploadParams.VOLUME);
                param.Add("@VOLUME_UM", _uploadParams.VOLUME_UM);
                param.Add("@WEIGHT", _uploadParams.WEIGHT);
                param.Add("@WEIGHT_UM", _uploadParams.WEIGHT_UM);
                param.Add("@FACTOR", _uploadParams.FACTOR);
                param.Add("@CASEQTY", _uploadParams.CASEQTY);
                param.Add("@USER_DEF1", _uploadParams.USER_DEF1);
                param.Add("@USER_DEF2", _uploadParams.DAIRYPACK);
                param.Add("@USER_DEF3", _uploadParams.TOTE);
                param.Add("@USER_DEF4", _uploadParams.USER_DEF4);
                param.Add("@USER_DEF5", _uploadParams.USER_DEF5);
                param.Add("@USER_DEF6", _uploadParams.USER_DEF6);
                param.Add("@USER_DEF7", _uploadParams.USER_DEF7);
                param.Add("@USER_DEF8", _uploadParams.USER_DEF8);
                param.Add("@CONTAINER", _uploadParams.CONTAINER);
                param.Add("@DAIRYPACK", _uploadParams.DAIRYPACK);
                param.Add("@TOTE", _uploadParams.TOTE);


                using (IDbConnection Conn = new SqlConnection(stringConn))
                {
                    var _id = Conn.Query<int>("CUBISCAN_UPLOAD_DIMENSIONS", param, null, true, null, CommandType.StoredProcedure);
                    _res = _id.FirstOrDefault();

                }
            }
            catch (Exception ex)
            {
                ErrorLogParams exparams = new ErrorLogParams();
                exparams.ITEM = xItem.ToString();
                exparams.DESCRIPTION = ex.Message.ToString();
                Update_ErrLog(exparams);
            }
            finally
            {

            }

            return _res;
        }

        public int Upload_Dimensions_ALT(UploadParams _uploadParams, string _mode)
        {
            int _res = 0;
            string xItem = "";

            _uploadParams.CONTAINER = _mode;
            if (_mode == "DAIRYPACK")
            {
                _uploadParams.LENGTH = 19.5M;
                _uploadParams.WIDTH = 12.5M;
                _uploadParams.HEIGHT = 9.0M;
                _uploadParams.QUANTITY_UM = "DP";
                _uploadParams.CONTAINER = "DP";
                _uploadParams.CASEQTY = _uploadParams.DAIRYPACK;
                _uploadParams.WEIGHT = (_uploadParams.WEIGHT * _uploadParams.DAIRYPACK);
            }
            else if (_mode == "Tote")
            {
                _uploadParams.LENGTH = 13.9M;
                _uploadParams.WIDTH = 20.75M;
                _uploadParams.HEIGHT = 12.45M;
                _uploadParams.QUANTITY_UM = "Tote";
                _uploadParams.CONTAINER = "Tote";
                _uploadParams.CASEQTY = _uploadParams.TOTE;
                _uploadParams.WEIGHT = (_uploadParams.WEIGHT * _uploadParams.TOTE);
            }

            try
            {
                DynamicParameters param = new DynamicParameters();
                xItem = _uploadParams.ITEM;
                param.Add("@ITEM", xItem);
                param.Add("@QUANTITY_UM", _uploadParams.QUANTITY_UM);
                param.Add("@COMPANY", _uploadParams.COMPANY);
                param.Add("@LENGTH", _uploadParams.LENGTH);
                param.Add("@WIDTH", _uploadParams.WIDTH);
                param.Add("@HEIGHT", _uploadParams.HEIGHT);
                param.Add("@DIMENSION_UM", _uploadParams.DIMENSION_UM);
                param.Add("@DIM_WEIGHT", _uploadParams.DIM_WEIGHT);
                param.Add("@VOLUME", _uploadParams.VOLUME);
                param.Add("@VOLUME_UM", _uploadParams.VOLUME_UM);
                param.Add("@WEIGHT", _uploadParams.WEIGHT);
                param.Add("@WEIGHT_UM", _uploadParams.WEIGHT_UM);
                param.Add("@FACTOR", _uploadParams.FACTOR);
                param.Add("@CASEQTY", _uploadParams.CASEQTY);
                param.Add("@USER_DEF1", _uploadParams.USER_DEF1);
                param.Add("@USER_DEF2", _uploadParams.DAIRYPACK);
                param.Add("@USER_DEF3", _uploadParams.TOTE);
                param.Add("@USER_DEF4", _uploadParams.USER_DEF4);
                param.Add("@USER_DEF5", _uploadParams.USER_DEF5);
                param.Add("@USER_DEF6", _uploadParams.USER_DEF6);
                param.Add("@USER_DEF7", _uploadParams.USER_DEF7);
                param.Add("@USER_DEF8", _uploadParams.USER_DEF8);
                param.Add("@CONTAINER", _uploadParams.CONTAINER);
                param.Add("@DAIRYPACK", _uploadParams.DAIRYPACK);
                param.Add("@TOTE", _uploadParams.TOTE);

                using (IDbConnection Conn = new SqlConnection(stringConn))
                {
                    var _id = Conn.Query<int>("CUBISCAN_UPLOAD_DIMENSIONS", param, null, true, null, CommandType.StoredProcedure);
                    _res = _id.FirstOrDefault();

                }
            }
            catch (Exception ex)
            {
                ErrorLogParams exparams = new ErrorLogParams();
                exparams.ITEM = xItem.ToString();
                exparams.DESCRIPTION = ex.Message.ToString();
                Update_ErrLog(exparams);
            }
            finally
            {

            }

            return _res;
        }

        public int Update_ErrLog(ErrorLogParams _errLogParams)
        {
            int _res = 0;

            try
            {
                DynamicParameters param = new DynamicParameters();
                param.Add("@ITEM", _errLogParams.ITEM);
                param.Add("@DESCRIPTION", _errLogParams.DESCRIPTION);
                param.Add("@FILENAME", _errLogParams.FILE_NAME);

                using (IDbConnection Conn = new SqlConnection(stringConn))
                {
                    var _id = Conn.Query<int>("CUBISCAN_UPDATE_ERROR_LOG", param, null, true, null, CommandType.StoredProcedure);
                    _res = _id.FirstOrDefault();

                }
            }
            catch (Exception ex)
            {

            }
            finally
            {

            }

            return _res;
        }

        public bool Update_SCALE()
        {
            bool _res = false;
            string xItem = "";
            int ID;

            try
            {
                // Get all none processed Items from CUBISCAN files ** \\wmsapp\ILS\Interface\Cubiscan
                ObservableCollection<UpdateParams> _upParams = new ObservableCollection<UpdateParams>();
                _upParams = GetAllUnprocessedItems(0);

                using (IDbConnection Conn = new SqlConnection(stringConn))
                {
                    foreach (UpdateParams p in _upParams)
                    {
                        DynamicParameters param = new DynamicParameters();
                        param.Add("@ITEM", p.ITEM);
                        param.Add("@CONTAINER", p.CONTAINER);
                        param.Add("@CASEQTY", p.CASEQTY);
                        param.Add("@LENGTH", p.LENGTH);
                        param.Add("@WIDTH", p.WIDTH);
                        param.Add("@HEIGHT", p.HEIGHT);
                        param.Add("@WEIGHT", p.WEIGHT);
                        param.Add("@USER_DEF1", p.USER_DEF1);
                        param.Add("@USER_DEF2", p.USER_DEF2);
                        param.Add("@USER_DEF3", p.USER_DEF3);
                        param.Add("@USER_DEF4", p.USER_DEF4);
                        param.Add("@USER_DEF5", p.USER_DEF5);
                        param.Add("@USER_DEF6", p.USER_DEF6);
                        param.Add("@USER_DEF7", p.USER_DEF7);
                        param.Add("@USER_DEF8", p.USER_DEF8);
                        param.Add("@TREAT_AS_LOOSE", IsTreatAsLoose(p.CONTAINER, p.ITEM));
                        param.Add("@COMPANY", p.COMPANY);

                        var _id = Conn.Query<int>("CUBISCAN_UPDATE_SCALE_UOM", param, null, true, null, CommandType.StoredProcedure);
                        ID = _id.FirstOrDefault();

                    }
                }
                _res = true;
            }
            catch (Exception ex)
            {
                ErrorLogParams exparams = new ErrorLogParams();
                exparams.ITEM = xItem.ToString();
                exparams.DESCRIPTION = ex.Message.ToString();
                exparams.FILE_NAME = "N/A";
                Update_ErrLog(exparams);
                _res = false;
            }
            finally
            {

            }

            return _res;
        }

        public string IsTreatAsLoose(string Container, string Item)
        {
            string _res = "Y";

            try
            {
                switch (Container)
                {
                    case "CS":
                        _res = "N";
                        break;
                    case "DP":
                        _res = "Y";
                        break;
                    case "EA":
                        _res = "Y";
                        break;
                    case "IP":
                        _res = "Y";
                        break;
                    case "PL":
                        _res = "N";
                        break;
                    case "TO":
                        _res = "Y";
                        break;
                    default:
                        _res = "Y";
                        break;
                }
            }
            catch (Exception ex)
            {
                ErrorLogParams exparams = new ErrorLogParams();
                exparams.ITEM = Item;
                exparams.DESCRIPTION = ex.Message.ToString();
                exparams.FILE_NAME = "N/A";
                Update_ErrLog(exparams);
            }

            return _res;
        }

        public bool PickUpFile()
        {
            bool _res = false;
            string _file = "";
            string[] FilePathSourceFile = Directory.GetFiles(ConfigurationManager.AppSettings["SOURCE"]);
            string FilePathDestinationFile = "";

            try
            {
                // Get all none processed Items from CUBISCAN files ** \\wmsapp\ILS\Interface\Cubiscan
                foreach (string file in FilePathSourceFile)
                {
                    _file = file.ToString();
                    if (PrepFile(_file))
                    {
                        FilePathDestinationFile = ConfigurationManager.AppSettings["PROCESSED"] + Path.GetFileName(_file);
                    }
                    else
                    {
                        FilePathDestinationFile = ConfigurationManager.AppSettings["ERROR"] + Path.GetFileName(_file);
                    }
                    File.Move(file, FilePathDestinationFile);
                }
                _res = true;
            }
            catch (Exception ex)
            {
                ErrorLogParams exparams = new ErrorLogParams();
                exparams.ITEM = "File";
                exparams.DESCRIPTION = ex.Message.ToString();
                exparams.FILE_NAME = _file.ToString();
                Update_ErrLog(exparams);
            }
            return _res;
        }

        public bool PrepFile(string FilePath)
        {
            string _DPorTOTE = "NONE";
            bool _res = false;
            string xItem = "";
            string xType = "EA";
            string xFile = FilePath;
            try
            {
                // Processed CUBISCAN files ** \\wmsapp\ILS\Interface\Cubiscan
                using (var reader = new StreamReader(FilePath))
                {
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(',');

                        if (values[0] != "Item number")
                        {
                            UploadParams _upParms = new UploadParams();
                            xItem = values[0];
                            xType = GetContainerType(xItem, FilePath);
                            if (xType != "EA")
                            { 
                                xItem = values[0].Substring(0,(values[0].Length - 3));
                            }
                            _upParms.ITEM = xItem;
                            _upParms.LENGTH = Convert.ToDecimal(values[3].ToString());
                            _upParms.WIDTH = Convert.ToDecimal(values[4].ToString());
                            _upParms.HEIGHT = Convert.ToDecimal(values[5].ToString());
                            _upParms.WEIGHT = Convert.ToDecimal(values[6].ToString());
                            _upParms.VOLUME = Convert.ToDecimal(values[7].ToString());
                            _upParms.DIM_WEIGHT = Convert.ToDecimal(values[8].ToString());
                            _upParms.DIMENSION_UM = values[9];
                            _upParms.WEIGHT_UM = values[10];
                            _upParms.VOLUME_UM = values[11];
                            _upParms.FACTOR = values[12].ToString();
                            _upParms.COMPANY = values[13];
                            if (!string.IsNullOrEmpty(values[15].ToString()))
                            {
                                _upParms.CASEQTY = Convert.ToInt32(values[15].ToString());
                            }
                            else
                            {
                                _upParms.CASEQTY = 0;
                            }
                            _upParms.QUANTITY_UM = xType;
                            _upParms.USER_DEF1 = "";
                            _upParms.USER_DEF2 = "";
                            _upParms.USER_DEF3 = "";

                            if (!string.IsNullOrEmpty(values[16].ToString()))
                            {
                                _upParms.DAIRYPACK = Convert.ToInt32(values[16].ToString());
                                _DPorTOTE = "DAIRYPACK";
                            }
                            else
                            {
                                _upParms.DAIRYPACK = 0;
                            }

                            if (!string.IsNullOrEmpty(values[17].ToString()))
                            {
                                _upParms.TOTE = Convert.ToInt32(values[17].ToString());
                                _DPorTOTE = "Tote"; 
                            }
                            else
                            {
                                _upParms.TOTE = 0;
                            }
                                                        
                            _upParms.USER_DEF4 = values[18];
                            _upParms.USER_DEF5 = values[19];
                            _upParms.USER_DEF6 = values[20];
                            if (values[21].ToString() == "") { _upParms.USER_DEF7 = 0; }
                            else { _upParms.USER_DEF7 = Convert.ToDecimal(values[21].ToString()); }
                            _upParms.USER_DEF8 = Convert.ToDecimal(values[22].ToString());
                            _upParms.CONTAINER = xType;

                            Upload_Dimensions(_upParms);
                            if (_DPorTOTE == "DAIRYPACK")
                            {
                                Upload_Dimensions_ALT(_upParms, _DPorTOTE);
                            }
                            else if (_DPorTOTE == "Tote")
                            {
                                Upload_Dimensions_ALT(_upParms, _DPorTOTE);
                            }
                            _res = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLogParams exparams = new ErrorLogParams();
                exparams.ITEM = xItem.ToString();
                exparams.DESCRIPTION = ex.Message.ToString();
                exparams.FILE_NAME = xFile.ToString();
                Update_ErrLog(exparams);
                _res = false;
            }
            return _res;
        }
        
        public string GetContainerType(string item, string FileName)
        {
            string _conttype = "EA";
            string _string = item.Substring((item.Length - 3), 3);

            try
            {
                switch (_string.ToUpper())
                {   // It does need to test for the dash (-) because that is the only way to know if it is container record
                    case "-CS":
                        _conttype = "CS";
                        break;
                    case "-DP":
                        _conttype = "DP";
                        break;
                    case "-EA":
                        _conttype = "EA";
                        break;
                    case "-IP":
                        _conttype = "IP";
                        break;
                    case "-PL":
                        _conttype = "PL";
                        break;
                    case "-TO":
                        _conttype = "Tote";
                        break;
                    default:
                        _conttype = "EA";
                        break;
                }
            }
            catch (Exception ex)
            {
                ErrorLogParams exparams = new ErrorLogParams();
                exparams.ITEM = item.ToString();
                exparams.DESCRIPTION = ex.Message.ToString();
                exparams.FILE_NAME = FileName.ToString();
                Update_ErrLog(exparams);
            }
            return _conttype;
        }

        public void SortBySequence(string item)
        {
            ObservableCollection<ItemSequence> itemlist = new ObservableCollection<ItemSequence>();
            ObservableCollection<SequenceSorter> updateList = new ObservableCollection<SequenceSorter>();

            List<string> scaleSEQ = new List<string>();
            scaleSEQ.Add("EA");
            scaleSEQ.Add("IP");
            scaleSEQ.Add("DP");
            scaleSEQ.Add("TO");
            scaleSEQ.Add("CS");
            scaleSEQ.Add("PL");

            try
            {
                DynamicParameters param = new DynamicParameters();
                param.Add("@ITEM", item);

                using (IDbConnection Conn = new SqlConnection(stringConn))
                {
                    var itemCol = Conn.Query<ItemSequence>("CUBISCAN_QUANTITY_UM_BY_ITEM", param, null, true, null, CommandType.StoredProcedure);
                    // GET ALL SEQUENCES BY ITEM
                    foreach (ItemSequence its in itemCol)
                    {
                        ItemSequence i = new ItemSequence();
                        i.ITEM = its.ITEM;
                        i.QUANTITY_UM = its.QUANTITY_UM;
                        i.SEQUENCE = its.SEQUENCE;
                        itemlist.Add(i);
                    }

                    // SORT ALL SEQUENCES BY SCALE
                    foreach (string seq in scaleSEQ)
                    {
                        SequenceSorter ss = new SequenceSorter();
                        ss.QUANTITY_UM = seq.ToString();
                        ss.IsPresent = false;

                        foreach (ItemSequence itm in itemlist)
                        {
                            if (itm.QUANTITY_UM == seq.ToString())
                            {
                                ss.IsPresent = true;
                            }
                        }
                        updateList.Add(ss);
                    }

                    //UPDATE ALL SEQUENCES ON SCALE
                    int xCounter = 1;
                    foreach (SequenceSorter x in updateList)
                    {
                        if (x.IsPresent == true)
                        {
                            DynamicParameters paramUpdate = new DynamicParameters();
                            paramUpdate.Add("@ITEM", item);
                            paramUpdate.Add("@CONTAINER", x.QUANTITY_UM);
                            paramUpdate.Add("@SEQUENCE", xCounter);

                            var _id = Conn.Query<int>("CUBISCAN_SORT_BY_QUANTITY_UM", paramUpdate, null, true, null, CommandType.StoredProcedure);
                            xCounter = xCounter + 1;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLogParams exparams = new ErrorLogParams();
                exparams.ITEM = item;
                exparams.DESCRIPTION = ex.Message.ToString();
                exparams.FILE_NAME = "N/A";
                Update_ErrLog(exparams);
            }
        }

        public ObservableCollection<UpdateParams> GetAllUnprocessedItems(int _mode)
        {
            // Get all none processed Items from CUBISCAN files ** \\wmsapp\ILS\Interface\Cubiscan
            ObservableCollection<UpdateParams> _upParams = new ObservableCollection<UpdateParams>();
            try
            {
                using (IDbConnection Conn = new SqlConnection(stringConn))
                {
                    DynamicParameters param = new DynamicParameters();
                    param.Add("@Mode", _mode);

                    var _UpCollection = Conn.Query<UpdateParams>("CUBISCAN_GET_UNPROCESSED_ITEMS", param, null, true, null, CommandType.StoredProcedure);
                    foreach (var _upC in _UpCollection)
                    {
                        UpdateParams _col = new UpdateParams();
                        _col.ITEM = _upC.ITEM;
                        if (_upC.CONTAINER != "EA")
                        {
                            _col.CASEQTY = _upC.CASEQTY;
                        }
                        else
                        {
                            _col.CASEQTY = 1;       // Just a double check, protecting the users from itself
                        }
                        _col.CONTAINER = _upC.CONTAINER;
                        _col.LENGTH = _upC.LENGTH;
                        _col.WIDTH = _upC.WIDTH;
                        _col.HEIGHT = _upC.HEIGHT;
                        _col.WEIGHT = _upC.WEIGHT;
                        _col.USER_DEF1 = _upC.USER_DEF1;
                        _col.USER_DEF2 = _upC.USER_DEF2;
                        _col.USER_DEF3 = _upC.USER_DEF3;
                        _col.USER_DEF4 = _upC.USER_DEF4;
                        _col.USER_DEF5 = _upC.USER_DEF5;
                        _col.USER_DEF6 = _upC.USER_DEF6;
                        _col.USER_DEF7 = _upC.USER_DEF7;
                        _col.USER_DEF8 = _upC.USER_DEF8;

                        _upParams.Add(_col);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLogParams exparams = new ErrorLogParams();
                exparams.ITEM = "SortBySequence";
                exparams.DESCRIPTION = ex.Message.ToString();
                exparams.FILE_NAME = "N/A";
                Update_ErrLog(exparams);
                return null;
            }
            finally
            {

            }
            return _upParams;
        }

        public void CleanUp()
        {
            string Item = "";

            try
            {
                // GET OPEN ITEMS
                ObservableCollection<UpdateParams> _upParams = new ObservableCollection<UpdateParams>();
                _upParams = GetAllUnprocessedItems(1);

                // UPDATE SEQUENCE
                foreach (UpdateParams pars in _upParams)
                {
                    Item = pars.ITEM;
                    SortBySequence(pars.ITEM);
                }

                using (IDbConnection Conn = new SqlConnection(stringConn))
                {
                    // CLOSE ITEMS
                    Conn.Query<int>("CUBISCAN_CLOSE_ITEMS", CommandType.StoredProcedure);
                }
            }
            catch (Exception ex)
            {
                ErrorLogParams exparams = new ErrorLogParams();
                exparams.ITEM = Item;
                exparams.DESCRIPTION = ex.Message.ToString();
                exparams.FILE_NAME = "N/A";
                Update_ErrLog(exparams);
            }
            finally
            {

            }
        }

    }
}
