protected void btnUpload_Click(object sender, EventArgs e)
    {
        try
        {
            if (FileUpload1.HasFile)
            {
                string FileName = Path.GetFileName(FileUpload1.PostedFile.FileName);
                string Extension = Path.GetExtension(FileUpload1.PostedFile.FileName);
                string FolderPath = ConfigurationManager.AppSettings["FolderPath"];

                string FilePath = Server.MapPath(FolderPath + FileName);
                FileUpload1.SaveAs(FilePath);
                currConnection2.Open();
                Import_To_Grid(FilePath, Extension);
                currConnection2.Close();
            }
        }
        catch (Exception ex)
        {
            lblMessage.Text = ex.Message;
        }
    }

    private void Import_To_Grid(string FilePath, string Extension)
    {
        string conStr = "";
        
        //verify file extension
        switch (Extension)
        {
            case ".xls": //Excel 97-03
                conStr = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                break;
            case ".xlsx": //Excel 07
                conStr = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                break;
        }
        conStr = String.Format(conStr, FilePath, "yes");
        OleDbConnection connExcel = new OleDbConnection(conStr);
        OleDbCommand cmdExcel = new OleDbCommand();
        OleDbDataAdapter oda = new OleDbDataAdapter();
        DataTable dt = new DataTable();
        cmdExcel.Connection = connExcel;

        //Get the name of First Sheet
        connExcel.Open();
        DataTable dtExcelSchema;
        dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
        connExcel.Close();

        //Read Data from First Sheet
        connExcel.Open();
        cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";
        oda.SelectCommand = cmdExcel;
        oda.Fill(dt);
        connExcel.Close();

        rpt.DataSource = dt;
        rpt.DataBind();
    }

    protected void btnSave_Click(object sender, EventArgs e)
    {
        lblError.Text = "";
        SqlConnection currConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["cesConnectionString"].ConnectionString);
        currConnection.Open();
        SqlCommand cmd = new SqlCommand();
        cmd.Connection = currConnection;
        cmd.CommandText = "insert into tblGosh (Gnumber,cityID) values (@Gnumber,'1')";

        SqlParameter p1 = new SqlParameter("@Gnumber", DbType.Int32);
        cmd.Parameters.Add(p1);

        try
        {
            foreach (RepeaterItem item in rpt.Items)
            {
                    p1.Value = ((Label)item.FindControl("lblNumber")).Text;
                    cmd.ExecuteNonQuery();
                }
        }
        catch (Exception ex)
        {
            lblError.Text = ex.Message;
        }

        finally
        {
            currConnection.Close();
        }
    }
