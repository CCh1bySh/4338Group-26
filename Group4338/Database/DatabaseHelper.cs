using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using Group4338.Models;

namespace Group4338.Database
{
    public class DatabaseHelper
    {
        private string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;Initial Catalog=EmployeesDB;Integrated Security=True";
        public void CreateEmployeesTable()
        {
            string query = @"
                IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Employees' AND xtype='U')
                CREATE TABLE Employees (
                    Id INT PRIMARY KEY,
                    Login NVARCHAR(50) NOT NULL,
                    Password NVARCHAR(50) NOT NULL,
                    Role NVARCHAR(50) NOT NULL
                )";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(query, conn);
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при создании таблицы: {ex.Message}");
                }
            }
        }

        public void InsertEmployee(Employee emp)
        {
            string query = "INSERT INTO Employees (Id, Login, Password, Role) VALUES (@Id, @Login, @Password, @Role)";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(query, conn);
                    cmd.Parameters.AddWithValue("@Id", emp.Id);
                    cmd.Parameters.AddWithValue("@Login", emp.Login);
                    cmd.Parameters.AddWithValue("@Password", emp.Password);
                    cmd.Parameters.AddWithValue("@Role", emp.Role);
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при вставке данных: {ex.Message}");
                }
            }
        }

        public void ClearEmployeesTable()
        {
            string query = "DELETE FROM Employees";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(query, conn);
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при очистке таблицы: {ex.Message}");
                }
            }
        }

        public List<Employee> GetAllEmployees()
        {
            List<Employee> employees = new List<Employee>();
            string query = "SELECT * FROM Employees ORDER BY Role, Login";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(query, conn);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    employees.Add(new Employee
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        Login = reader["Login"].ToString(),
                        Password = reader["Password"].ToString(),
                        Role = reader["Role"].ToString()
                    });
                }
            }

            return employees;
        }

        public List<string> GetDistinctRoles()
        {
            List<string> roles = new List<string>();
            string query = "SELECT DISTINCT Role FROM Employees ORDER BY Role";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(query, conn);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    roles.Add(reader["Role"].ToString());
                }
            }

            return roles;
        }

        public List<Employee> GetEmployeesByRole(string role)
        {
            List<Employee> employees = new List<Employee>();
            string query = "SELECT * FROM Employees WHERE Role = @Role ORDER BY Login";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@Role", role);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    employees.Add(new Employee
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        Login = reader["Login"].ToString(),
                        Password = reader["Password"].ToString(),
                        Role = reader["Role"].ToString()
                    });
                }
            }

            return employees;
        }
    }
}