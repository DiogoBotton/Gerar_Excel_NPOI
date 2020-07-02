using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExportarArqExcel.Models
{
    public class Usuarios
    {
        public static List<Usuarios> listUsuarios = new List<Usuarios>();
        public long Id { get; set; }
        public string Nome { get; set; }
        public string Email { get; set; }
        public string Senha { get; set; }
        public string Cpf { get; set; }

        public Usuarios()
        {

        }

        public Usuarios(string nome, string email, string senha, string cpf)
        {
            Nome = nome;
            Email = email;
            Senha = senha;
            Cpf = cpf;
        }
    }
}
