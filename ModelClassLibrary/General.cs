using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Reflection;

public static class General
{
    static string strConnectionString = "Data Source=10.250.19.38;Initial Catalog=SP_PRICING;User ID=payrollgiving;Password=payrollgiving;";   //Dev
    //static string strConnectionString = "Data Source=10.250.19.38;Initial Catalog=SP_PRICING_UAT;User ID=payrollgiving;Password=payrollgiving;";   //UAT
    //static string strConnectionString = "Data Source=edemumkaldbs002;Initial Catalog=SP_PRICING_Live;User ID=SPPricing;Password=sppricing~;";     //LIVE

    public static DataSet ExecuteDataSet(string strProcedureName)
    {
        SqlConnection objConn = new SqlConnection(strConnectionString);
        SqlCommand cmd = new SqlCommand();

        cmd.CommandType = CommandType.StoredProcedure;
        cmd.CommandText = strProcedureName;
        cmd.Connection = objConn;
        cmd.CommandTimeout = 0;

        objConn.Open();

        SqlDataAdapter objSqlDataAdapter = new SqlDataAdapter(cmd);
        DataSet dsResult = new DataSet();
        objSqlDataAdapter.Fill(dsResult);

        if (objConn.State == ConnectionState.Open)
            objConn.Close();

        return dsResult;
    }

    public static DataSet ExecuteDataSet(string strProcedureName, params object[] cmdParams)
    {
        SqlConnection objConn = new SqlConnection(strConnectionString);
        SqlCommand cmd = new SqlCommand();

        cmd.CommandType = CommandType.StoredProcedure;
        cmd.CommandText = strProcedureName;
        cmd.Connection = objConn;
        cmd.CommandTimeout = 0;

        objConn.Open();
        SqlCommandBuilder.DeriveParameters(cmd);

        for (int i = 0; i < cmdParams.Length; i++)
        {
            cmd.Parameters[i + 1].Value = cmdParams[i];
        }

        SqlDataAdapter objSqlDataAdapter = new SqlDataAdapter(cmd);
        DataSet dsResult = new DataSet();
        objSqlDataAdapter.Fill(dsResult);

        if (objConn.State == ConnectionState.Open)
            objConn.Close();

        return dsResult;
    }

    public static void ReflectSingleData(object objSource, object drRow)
    {
        Type typeSource = objSource.GetType();
        PropertyInfo[] propSource = typeSource.GetProperties();

        Type typeDestination = drRow.GetType();
        
        if (propSource != null)
        {
            foreach (PropertyInfo propInfo in propSource)
            {
                PropertyInfo propDestination = typeDestination.GetProperty(propInfo.Name);          

                if (propDestination != null)
                {
                    //object value = propDestination.GetValue(drRow, null);                        

                    switch (propInfo.PropertyType.FullName.ToLower())
                    {
                        case "system.string":
                            propInfo.SetValue(objSource, Convert.ToString(propDestination.GetValue(drRow, null)), null);
                            break;
                        case "system.int32":
                            propInfo.SetValue(objSource, Convert.ToInt32(propDestination.GetValue(drRow, null)), null);
                            break;
                        case "system.int64":
                            propInfo.SetValue(objSource, Convert.ToInt64(propDestination.GetValue(drRow, null)), null);
                            break;
                        case "system.decimal":
                            propInfo.SetValue(objSource, Convert.ToDecimal(propDestination.GetValue(drRow, null)), null);
                            break;
                        case "system.bool":
                            propInfo.SetValue(objSource, Convert.ToBoolean(propDestination.GetValue(drRow, null)), null);
                            break;
                        case "system.double":
                            propInfo.SetValue(objSource, Convert.ToDouble(propDestination.GetValue(drRow, null)), null);
                            break;
                        case "system.datetime":
                            propInfo.SetValue(objSource, Convert.ToDateTime(propDestination.GetValue(drRow, null)), null);
                            break;
                        default:
                            break;
                    }
                }
            }
        }
    }

    public static List<object> ReflectArray(List<object> SourceList, object Destination)
    {
        List<object> DestinationList = new List<object>();

        foreach (var Source in SourceList)
        {
            General.ReflectSingleData(Destination, SourceList);
            DestinationList.Add(Destination);
        }

        return DestinationList;
    }

    public static void LogError(string strErrorDescription, string strStackTrace, string strClassName, string strMethodName)
    {
        DataSet dsResult = ExecuteDataSet("SP_LOG_ERROR", strErrorDescription, strStackTrace, strClassName, strMethodName);
    }

    public static DataSet GetLatestSpot()
    {
        DataSet dsResult = General.ExecuteDataSet("SP_GET_LATEST_SPOT");

        return dsResult;
    }
}