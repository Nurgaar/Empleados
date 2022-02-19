using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Empleados
{
    public partial class FrmEmpleados : Form
    {
        IDbConnection Conexiondb;
        ListBox[] LST = new ListBox[5];

        public FrmEmpleados()
        {
            InitializeComponent();
            LST[0] = LstApellidos;
            LST[1] = LstComision;
            LST[2] = LstSalario;
            LST[3] = LstOficio;
            LST[4] = LstFechaAlta;
        }

        private void ConexionBD(object sender, EventArgs e)
        {

            try
            {
                string cadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =C:\\Users\\usuario\\OneDrive - IES Tetuán\\Escritorio\\temp\\Emple.mdb";
                Conexiondb = new OleDbConnection(cadenaConexion);
                Conexiondb.Open();
                MessageBox.Show("Establecida conexión con la base de datos", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);




                //creo y ejecuto el comando con una sentencia sql
               

                IDbCommand comand = Conexiondb.CreateCommand();
                comand.CommandText = "Select * from DEPART";
                IDataReader leer = comand.ExecuteReader();

                //Que me muestre toda la info de la tabla depart y me la añada en ambas listbox
                //partiendo de la sentencia SQL
                while (leer.Read())
                {
                    LstIdDepart.Items.Add(Convert.ToInt32(leer.GetValue(0)));
                    LstDepartamento.Items.Add(leer.GetValue(1).ToString());
                    LstLocalizacion.Items.Add(leer.GetValue(2).ToString());

                }
            }

            catch (Exception err)
            {
                MessageBox.Show(err.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            

        }

        private void CerrarBD(object sender, FormClosingEventArgs e)
        {
            if (Conexiondb.State == ConnectionState.Open)
            {
                MessageBox.Show("Cerrando la conexión con la BD", "Información");
                Conexiondb.Close();
            }
        }

        private void BtnSalir_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void DepartamentoElegido(object sender, EventArgs e)
        {
            limpiarLST();

            int i;

            if (sender == LstDepartamento)
            {
                i = LstDepartamento.SelectedIndex;
                LstLocalizacion.SelectedIndex = i;


            }
            else
            {
                i = LstLocalizacion.SelectedIndex;
                LstDepartamento.SelectedIndex = i;
            }

            string id = LstIdDepart.Items[i].ToString();

            IDbCommand comando = Conexiondb.CreateCommand();

            comando.CommandText = "select * from emple where dept_no = " + id;

            IDataReader leer = comando.ExecuteReader();


            while (leer.Read())
            {
                LstIdEmple.Items.Add(leer.GetValue(0).ToString());
                LstApellidos.Items.Add(leer.GetValue(1).ToString());
                LstOficio.Items.Add(leer.GetValue(2).ToString());
                LstSalario.Items.Add(leer.GetValue(5).ToString());
                LstFechaAlta.Items.Add(leer.GetValue(4).ToString());
                LstComision.Items.Add(leer.GetValue(6).ToString());

            }

            leer.Close();

        }
     
        private void Borrar(object sender, EventArgs e)
        {
            // Eliminar un empleado ya existente en la BD
            

            int i = LstApellidos.SelectedIndex;
          if (i != -1)
            {
                String claveEmpleado = LstIdEmple.Items[i].ToString();


                IDbCommand comando = Conexiondb.CreateCommand();

                comando.CommandText = "delete from emple where emp_no = " + claveEmpleado;

                comando.ExecuteNonQuery();


            }
            else
                MessageBox.Show("Selecciona un empleado antes de pulsar el botón");

            limpiarLST();
            vaciarTXT();
        }

        private void Actualizar(object sender, EventArgs e)
        {
            // Actualizar la información de un empleado ya existente en la BD. 
           

            OleDbCommand comando = (OleDbCommand)Conexiondb.CreateCommand();

            comando.CommandText = "update emple set oficio = @ofi, salario = @salar, comision = @comis where emp_no =" + LstIdEmple.Items[LstApellidos.SelectedIndex];

            comando.Parameters.AddWithValue("@ofi", TxtOficio.Text);
            comando.Parameters.AddWithValue("@salar", Convert.ToDecimal(TxtSalario.Text));
            comando.Parameters.AddWithValue("@comis", Convert.ToDecimal(TxtComision.Text));


            int i = comando.ExecuteNonQuery();

            if (i > 0)
            {
                MessageBox.Show("Actualización hecha correctamente");

            }
            else
                MessageBox.Show("No se ha podido actualizar correctamente");


            limpiarLST();
            vaciarTXT();

        }

        private void Nuevo(object sender, EventArgs e)
        {
            //Se da de Alta un nuevo empleado en la BD

            //He puesto el TextBox de Fecha de Alta en ReadOnly=True ya que se va a poner la fecha de hoy automáticamente

            OleDbCommand comando = (OleDbCommand)Conexiondb.CreateCommand();


            comando.CommandText = "select max(emp_no) from emple";
            int max = Convert.ToInt32(comando.ExecuteScalar());



            comando = (OleDbCommand)Conexiondb.CreateCommand();


            if (LstDepartamento.SelectedIndex != -1)
            {
                comando.CommandText = "insert into emple values (@emp_no, @apellido, @oficio, @direc, @fecha, @salar, @comis, @depto)";

                comando.Parameters.AddWithValue("@emp_no", OleDbType.Numeric).Value = max + 1;
                comando.Parameters.AddWithValue("@apellido", TxtApellido.Text);
                comando.Parameters.AddWithValue("@oficio", OleDbType.VarChar).Value = TxtOficio.Text;
                comando.Parameters.AddWithValue("@direc", OleDbType.Numeric).Value = 0;
                comando.Parameters.AddWithValue("@fecha", OleDbType.DBDate).Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
                comando.Parameters.AddWithValue("@salar", OleDbType.Decimal).Value = TxtSalario.Text;
                comando.Parameters.AddWithValue("@comis", OleDbType.Decimal).Value = TxtComision.Text;
                comando.Parameters.AddWithValue("@depto", OleDbType.Decimal).Value = LstIdDepart.Items[LstDepartamento.SelectedIndex];

                comando.ExecuteNonQuery();
                vaciarTXT();

            }
            else
                MessageBox.Show("Selecciona un departamento para poder introducir el empleado correctamente");

            limpiarLST();


        }

        private void BuscarEmpleado(object sender, EventArgs e)
        {
            //Busco el empleado que quiero introduciendolo en el text box y lo muestro en las lisbox

            OleDbCommand comando = (OleDbCommand)Conexiondb.CreateCommand();


            if (Lstbuscarpor.SelectedItem != null)
            {
                string campo_Seleccionado = Lstbuscarpor.SelectedItem.ToString();

               limpiarLST();


                comando.CommandType = CommandType.Text;

                comando.CommandText = String.Format("select * from emple where {0} = @apellido", campo_Seleccionado);


                comando.Parameters.AddWithValue("@apellido", Txtbuscar.Text.ToUpper());


                IDataReader leer = comando.ExecuteReader();


                while (leer.Read())
                {
                    LstApellidos.Items.Add(leer.GetValue(1).ToString());
                    LstOficio.Items.Add(leer.GetValue(2).ToString());
                    LstFechaAlta.Items.Add(leer.GetValue(4).ToString());
                    LstSalario.Items.Add(leer.GetValue(5).ToString());
                    LstComision.Items.Add(leer.GetValue(6).ToString());
                }

                leer.Close();


            }
            else            
                MessageBox.Show("Selecciona un campo de búsqueda","ERROR", MessageBoxButtons.OK);
            

        }
              

        private void Btn_Abajo(object sender, EventArgs e)
        {
            foreach (ListBox box in LST)
            {
                if (box.Items.Count - 1 == box.SelectedIndex || box.SelectedIndex == -1)
                {
                    box.SelectedIndex = 0;
                }
                else
                 {
                    box.SelectedIndex++;

                }
            }
        }

        private void Btn_Arriba(object sender, EventArgs e)
        {
            foreach (ListBox box in LST)
            {
                if (box.SelectedIndex == 0 || box.SelectedIndex == -1)
                {
                    box.SelectedIndex = box.Items.Count - 1;
                }
                else
                {
                    box.SelectedIndex--;
                }
            }
        }

        private void limpiarLST()
        {
            LstApellidos.Items.Clear();
            LstOficio.Items.Clear();
            LstSalario.Items.Clear();
            LstFechaAlta.Items.Clear();
            LstComision.Items.Clear();
            LstIdEmple.Items.Clear();


        }

        private void vaciarTXT()
        {
            TxtApellido.Text = "";
            TxtComision.Text = "";
            TxtOficio.Text = "";
            TxtSalario.Text = "";
            TxtFechaAlta.Text = "";
        }

        private void LstSeleccionada(int i)
        {

            //selecciona todas las ListBox, cuando haces un solo SelectIndex, luego se llama en cada evento click de las LST
            foreach (ListBox caja in LST)
            {
                caja.SelectedIndex = i;
            }

        }

        private void RellenarTxt(int i)
        {
            TxtApellido.Text = LstApellidos.Items[i].ToString();
            TxtOficio.Text = LstOficio.Items[i].ToString();
            TxtComision.Text = LstComision.Items[i].ToString();
            TxtFechaAlta.Text = LstFechaAlta.Items[i].ToString();
            TxtSalario.Text = LstSalario.Items[i].ToString();
        }

        private void apellidos_click(object sender, EventArgs e)
        {
            int i = LstApellidos.SelectedIndex;
            RellenarTxt(i);
            LstSeleccionada(i);
        }

        private void oficio_click(object sender, EventArgs e)
        {
            int i = LstOficio.SelectedIndex;
            RellenarTxt(i);
            LstSeleccionada(i);
        }

        private void salario_click(object sender, EventArgs e)
        {
            int i = LstSalario.SelectedIndex;
            RellenarTxt(i);
            LstSeleccionada(i);
        }

        private void fecha_click(object sender, EventArgs e)
        {
            int i = LstFechaAlta.SelectedIndex;
            RellenarTxt(i);
            LstSeleccionada(i);
        }

        private void comision_click(object sender, EventArgs e)
        {
            int i = LstComision.SelectedIndex;
            RellenarTxt(i);
            LstSeleccionada(i);
        }
    }

}
