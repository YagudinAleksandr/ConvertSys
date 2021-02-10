using System;
using System.Windows.Forms;

namespace ConvertSys
{
    public class CreateAccessDB
    {
        public static void CreateNiewAccessDatabase()
        {

            string fileName = "NewDB.mdb";


            ADOX.Catalog cat = new ADOX.Catalog();


            try
            {
                //Создаем базу данных
                cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + fileName + "; Jet OLEDB:Engine Type=5");

                //Закрываем базу данных
                ADODB.Connection con = cat.ActiveConnection as ADODB.Connection;
                if (con != null)
                    con.Close();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            cat = null;

            //Строка подключения к созданной базе данных
            string connectionString = string.Format("Provider={0}; Data Source={1}; Jet OLEDB:Engine Type={2}", "Microsoft.Jet.OLEDB.4.0", "NewDB.mdb", 5);

            using (var con = new System.Data.OleDb.OleDbConnection(connectionString))
            {
                con.Open();


                //Создание таблиц баз данных
                System.Data.OleDb.OleDbCommand oleDbCommand = new System.Data.OleDb.OleDbCommand();
                oleDbCommand.Connection = con;


                try
                {
                    oleDbCommand.CommandText = @"CREATE TABLE Parametry(" +
                        "Parametr Memo WITH COMP," +
                        "ParamGrp Text(32) WITH COMP NOT NULL," +
                        "ParamImia Text(32) WITH COMP NOT NULL" +
                    ");";
                    oleDbCommand.ExecuteNonQuery();


                    oleDbCommand.CommandText = @"CREATE TABLE TblKvr(" +
                        "NomZ AutoIncrement," +
                        "KvrNomK Long NOT NULL," +
                        "KvrNomKD Text(4) DEFAULT ''," +
                        "KvrPls Decimal(16,4) NOT NULL DEFAULT 0," +
                        "Relief SmallInt NOT NULL DEFAULT 0," +
                        "PozharKlsKvr Decimal(3,1) DEFAULT 0," +
                        "LupINN Long NOT NULL DEFAULT 0," +
                        "LupTaksator Long NOT NULL DEFAULT 0," +
                        "LesBaz Long NOT NULL DEFAULT 0," +
                        "Plansh Long NOT NULL DEFAULT 0," +
                        "TehUchas Long DEFAULT 0," +
                        "Obhod Long DEFAULT 0," +
                        "Info Memo DEFAULT ''," +
                        "GodLu Long NOT NULL DEFAULT 0," +
                        "RajonTaks Long NOT NULL DEFAULT 0," +
                        "RazTaks Long NOT NULL DEFAULT 0," +
                        "CONSTRAINT PrimaryKey PRIMARY KEY(NomZ)" +
                     ");";
                    oleDbCommand.ExecuteNonQuery();

                    oleDbCommand.CommandText = @"CREATE TABLE TblVyd(" +
                        "NomZ AutoIncrement," +
                        "NomSoed Long NOT NULL," +
                        "Vybor0 SmallInt NOT NULL DEFAULT 0," +
                        "Vybor1 SmallInt NOT NULL DEFAULT 0," +
                        "KvrNom SmallInt NOT NULL," +
                        "KvrNomD Text(4) DEFAULT ''," +
                        "VydNom SmallInt NOT NULL," +
                        "VydNomD Text(4) DEFAULT ''," +
                        "VydPls Decimal(16,4) NOT NULL DEFAULT 0," +
                        "VydPlsFiks Bit NOT NULL DEFAULT False," +
                        "VydTip SmallInt NOT NULL DEFAULT 0," +
                        "KatZem Long NOT NULL DEFAULT 0," +
                        "KatZasch Long NOT NULL DEFAULT 0," +
                        "FunZona Long NOT NULL DEFAULT 0," +
                        "HozKat Long NOT NULL DEFAULT 0," +
                        "OZU Long NOT NULL DEFAULT 0," +
                        "HozSek Long NOT NULL DEFAULT 0," +
                        "PorodaPrb Long NOT NULL DEFAULT 0," +
                        "PorodaCel Long NOT NULL DEFAULT 0," +
                        "Bonitet SmallInt NOT NULL DEFAULT 0," +
                        "TipLesa Long NOT NULL DEFAULT 0," +
                        "TLU Long NOT NULL DEFAULT 0," +
                        "ZapasVyd Decimal(10,1) NOT NULL DEFAULT 0," +
                        "ZapasZah Decimal(5,1) NOT NULL DEFAULT 0," +
                        "ZapasZahL Decimal(5,1) NOT NULL DEFAULT 0," +
                        "ZapasSuh Decimal(5,1) NOT NULL DEFAULT 0," +
                        "VozGrpVyd SmallInt NOT NULL DEFAULT 0," +
                        "VozKls SmallInt NOT NULL DEFAULT 0," +
                        "VozRubki SmallInt NOT NULL DEFAULT 0," +
                        "PozharKlsVyd Decimal(3,1) NOT NULL DEFAULT 0," +
                        "GodAkt DateTime NOT NULL DEFAULT #1/1/1999#," +
                        "MetodTaks SmallInt NOT NULL DEFAULT 0," +
                        "RajonMunic SmallInt NOT NULL DEFAULT 0," +
                        "SklonEkspoz SmallInt NOT NULL DEFAULT 0," +
                        "SklonKrut SmallInt NOT NULL DEFAULT 0," +
                        "VNUM SmallInt NOT NULL DEFAULT 0," +
                        "SklonEroz SmallInt NOT NULL DEFAULT 0," +
                        "SklonErozStep SmallInt NOT NULL DEFAULT 0," +
                        "Strata SmallInt NOT NULL DEFAULT 0," +
                        "Info Memo DEFAULT ''," +
                        "DataIzm DateTime," +
                        "LesohozZona Long NOT NULL DEFAULT 0," +
                        "LesVosTip Long NOT NULL DEFAULT 0," +
                        "CONSTRAINT PrimaryKey PRIMARY KEY(NomZ)" +
                    ");";
                    oleDbCommand.ExecuteNonQuery();

                    oleDbCommand.CommandText = @"CREATE TABLE TblVydDopMaket(" +
                            "Maket SmallInt NOT NULL DEFAULT 0," +
                            "NomSoed Long NOT NULL," +
                            "NomZ AutoIncrement," +
                            "Vybor0 SmallInt DEFAULT 0," +
                            "CONSTRAINT PrimaryKey PRIMARY KEY(NomZ)" +
                        ");";
                    oleDbCommand.ExecuteNonQuery();

                    oleDbCommand.CommandText = @"CREATE TABLE TblVydDopParam(" +
                            "NomSoed Long NOT NULL," +
                            "NomZ AutoIncrement," +
                            "Parametr Text(32) NOT NULL DEFAULT ''," +
                            "ParamId SmallInt NOT NULL DEFAULT 0," +
                            "CONSTRAINT PrimaryKey PRIMARY KEY(NomZ)" +
                        ");";
                    oleDbCommand.ExecuteNonQuery();

                    oleDbCommand.CommandText = @"CREATE TABLE TblOshibki(" +
                            "IndZnaka SmallInt DEFAULT 0," +
                            "NomSoed Long NOT NULL," +
                            "NomZ AutoIncrement," +
                            "Oshibka Long NOT NULL DEFAULT 0," +
                            "CONSTRAINT PrimaryKey PRIMARY KEY(NomZ)" +
                        ");";
                    oleDbCommand.ExecuteNonQuery();

                    oleDbCommand.CommandText = @"CREATE TABLE TblVydIarus(" +
                        "NomZ AutoIncrement," +
                        "NomSoed Long NOT NULL," +
                        "Vybor0 SmallInt DEFAULT 0," +
                        "Iarus SmallInt NOT NULL DEFAULT 0," +
                        "Sostav Text(48) DEFAULT ''," +
                        "VozrastIar SmallInt NOT NULL DEFAULT 0," +
                        "VozGrpIar SmallInt NOT NULL DEFAULT 0," +
                        "VysotaIar Decimal(5,1) NOT NULL DEFAULT 0," +
                        "DiamIar SmallInt NOT NULL DEFAULT 0," +
                        "Polnota Decimal(3,1) NOT NULL DEFAULT 0," +
                        "SumPlsS Decimal(4,1) DEFAULT 0," +
                        "SumPlsS_A Decimal(4,1) DEFAULT 0," +
                        "ZapasGa Decimal(6,1) NOT NULL DEFAULT 0," +
                        "Prois SmallInt NOT NULL DEFAULT 0," +
                        "IGP Bit NOT NULL," +
                        "KolStvol Decimal(4, 1) NOT NULL DEFAULT 0," +
                        "ProcentPrizh SmallInt NOT NULL DEFAULT 0," +
                        "Gustota SmallInt NOT NULL DEFAULT 0," +
                        "IarusNom Long NOT NULL DEFAULT 0," +
                        "Ocenka SmallInt NOT NULL DEFAULT 0," +
                        "CONSTRAINT PrimaryKey PRIMARY KEY(NomZ)" +
                    ");";
                    oleDbCommand.ExecuteNonQuery();

                    oleDbCommand.CommandText = @"CREATE TABLE TblVydMer(" +
                        "Info Memo DEFAULT ''," +
                        "MerKl Long NOT NULL DEFAULT 0," +
                        "MerNom SmallInt NOT NULL," +
                        "MerProcent SmallInt NOT NULL DEFAULT 0," +
                        "MerRTK Long NOT NULL DEFAULT 0," +
                        "NomSoed Long NOT NULL," +
                        "NomZ AutoIncrement," +
                        "Vybor0 SmallInt NOT NULL DEFAULT 0," +
                        "CONSTRAINT PrimaryKey PRIMARY KEY(NomZ)" +
                    ");";
                    oleDbCommand.ExecuteNonQuery();

                    oleDbCommand.CommandText = @"CREATE TABLE TblVydPoroda(" +
                        "Diam_A Decimal(5,1) DEFAULT 0," +
                        "DiamPor SmallInt DEFAULT 0," +
                        "DiamTaks SmallInt DEFAULT 0," +
                        "KlsTov SmallInt NOT NULL DEFAULT 0," +
                        "KoefSos SmallInt NOT NULL DEFAULT 0," +
                        "MerProcent SmallInt DEFAULT 0," +
                        "NomSoed Long NOT NULL," +
                        "NomZ AutoIncrement," +
                        "Poroda Long NOT NULL DEFAULT 0," +
                        "PorodaNom SmallInt NOT NULL," +
                        "ProisPor Long NOT NULL DEFAULT 0," +
                        "Vozrast_A SmallInt DEFAULT 0," +
                        "VozrastPor SmallInt NOT NULL DEFAULT 0," +
                        "VozrastTaks SmallInt DEFAULT 0," +
                        "Vybor0 SmallInt DEFAULT 0," +
                        "Vysota_A Decimal(5,1) DEFAULT 0," +
                        "VysotaPor Decimal(5,1) NOT NULL DEFAULT 0," +
                        "VysotaTaks SmallInt DEFAULT 0," +
                        "CONSTRAINT PrimaryKey PRIMARY KEY(NomZ)" +
                    ");";

                    oleDbCommand.ExecuteNonQuery();
                    //Выстраиваем связи базы данных
                    oleDbCommand.CommandText = @"ALTER TABLE TblOshibki " +
                        "ADD CONSTRAINT TblVydTblOshibki FOREIGN KEY (NomSoed) REFERENCES TblVyd ON UPDATE CASCADE ON DELETE CASCADE";
                    oleDbCommand.ExecuteNonQuery();

                    oleDbCommand.CommandText = @"ALTER TABLE TblVyd " +
                        "ADD CONSTRAINT TblVyd_TblKv_FK FOREIGN KEY (NomSoed) REFERENCES TblKvr ON UPDATE CASCADE ON DELETE CASCADE;";
                    oleDbCommand.ExecuteNonQuery();

                    oleDbCommand.CommandText = @"ALTER TABLE TblVydDopMaket " +
                        "ADD CONSTRAINT TblVydTblVydDopMak FOREIGN KEY (NomSoed) REFERENCES TblVyd ON UPDATE CASCADE ON DELETE CASCADE";
                    oleDbCommand.ExecuteNonQuery();

                    oleDbCommand.CommandText = @"ALTER TABLE TblVydDopParam " +
                        "ADD CONSTRAINT Tblam FOREIGN KEY (NomSoed) REFERENCES TblVydDopMaket ON UPDATE CASCADE ON DELETE CASCADE;";
                    oleDbCommand.ExecuteNonQuery();

                    oleDbCommand.CommandText = @"ALTER TABLE TblVydIarus " +
                        "ADD CONSTRAINT TblFK FOREIGN KEY (NomSoed) REFERENCES TblVyd ON UPDATE CASCADE ON DELETE CASCADE;";
                    oleDbCommand.ExecuteNonQuery();

                    oleDbCommand.CommandText = @"ALTER TABLE TblVydMer " +
                        "ADD CONSTRAINT TblVydMer_TblVyd_FK FOREIGN KEY (NomSoed) REFERENCES TblVyd ON UPDATE CASCADE ON DELETE CASCADE;";
                    oleDbCommand.ExecuteNonQuery();

                    oleDbCommand.CommandText = @"ALTER TABLE TblVydPoroda " +
                        "ADD CONSTRAINT TblVydYarusSostaw_TblVydYarus_FK FOREIGN KEY (NomSoed) REFERENCES TblVydIarus ON UPDATE CASCADE ON DELETE CASCADE;";
                    oleDbCommand.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                    return;
                }
                finally
                {
                    //Закрываем соединение с базой данных
                    oleDbCommand.Connection.Close();
                }


                MessageBox.Show("База данных успешно создана!");


            }

        }
    }
}
