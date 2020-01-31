using System;
using CubiscanInterface.DBHelpers;
using CubiscanInterface.Models;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using Dapper;

namespace CubiscanInterface.DBHelpers
{
    class HPBWorkBench : IDisposable
    {
        #region "Dispose Logic"
        private bool disposed = false;
        public string stringConn;

        /// <summary>
        /// Constructor
        /// </summary>
        public HPBWorkBench()
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

        //public ObservableCollection<Items> GetItems(string itm)
        //{
        //    ObservableCollection<Items> dep = new ObservableCollection<Items>();

        //    DynamicParameters param = new DynamicParameters();
        //    param.Add("@ITEM", itm);

        //    using (IDbConnection Conn = new SqlConnection(stringConn))
        //    {
        //        var dep2 = Conn.Query<Items>("DASH_GETITEM", param, null, true, null, CommandType.StoredProcedure);

        //        foreach (var user in dep2)
        //        {
        //            Items item = new Items();
        //            item.INTERNAL_ITEM_NUM = user.INTERNAL_ITEM_NUM;
        //            item.ITEM = user.ITEM;
        //            item.COMPANY = user.COMPANY;
        //            item.DESCRIPTION = user.DESCRIPTION;
        //            item.ITEM_CATEGORY1 = user.ITEM_CATEGORY1;
        //            item.DATE_TIME_STAMP = user.DATE_TIME_STAMP;
        //            dep.Add(item);
        //        }
        //    }

        //    return dep;
        //}

    }
}
