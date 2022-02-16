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
        public FrmEmpleados()
        {
            InitializeComponent();
        }

        private void ConexionBD(object sender, EventArgs e)
        {

            try
            {
                string cadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\\temp\\Emple.mdb";
                Conexiondb = new OleDbConnection(cadenaConexion);
                Conexiondb.Open();




                //creo y ejecuto el comando con una sentencia sql

                IDbCommand comand = Conexiondb.CreateCommand();
                comand.CommandText = "Select * from DEPART";
                IDataReader leer = comand.ExecuteReader();

                while (leer.Read())
                {
                    LstDepartamento.Items.Add(Convert.ToInt32(leer.GetValue(0)));
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

            string clave = LstIdDepart.Items[i].ToString();

            //se hace otra conexión, pero esta vez con la consulta buscadno los empleados dependiendo de su id de departamento al que pertenecen
            IDbCommand comand = Conexiondb.CreateCommand();
            comand.CommandText = "select * from EMPLE where dept_no= " + clave;
            IDataReader leer=comand.ExecuteReader();

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
            // Eliminar la información de un empleado ya existente en la BD
            //No se puede borrar a un empleado que no existe, por lo que se avisa con MessageBox
        }

        private void Guardar(object sender, EventArgs e)
        {
            // Actualizar la información de un empleado ya existente en la BD. 
            //No se puede actualizar un empleado que no existe, por lo que se avisa con MessageBox

        }

        private void Nuevo(object sender, EventArgs e)
        {
            //Se da de Alta un nuevo empleado en la BD
        }

        private void BuscarEmpleado(object sender, EventArgs e)
        {
            //Busco el empleado que quiero introduciendolo en el text box y lo muestro en las lisbox
        }

     
    }

}
