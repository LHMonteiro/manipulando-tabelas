using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Polimorfico
{
    public class Descricao
    {
        [Description("Codigo")]
        public string Codigo { get; set; }

        [Description("Editora")]
        public string Editora { get; set; }

        [Description("Genero")]
        public string Genero { get; set; }

        [Description("Livre")]
        public string Livro { get; set; }

        [Description("Autor")]
        public string Autor { get; set; }

        
        public Descricao(string codigo,string editora, string genero, string livro, string autor){
            Codigo = codigo;
            Editora = editora;
            Genero = genero;
            Livro = livro;
            Autor = autor;
        }
    }
}