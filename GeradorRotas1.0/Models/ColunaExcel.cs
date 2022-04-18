using Microsoft.AspNetCore.Mvc.Rendering;
using System.Collections.Generic;

namespace GeradorRotas.Models
{
    public class ColunaExcel
    {
        public IEnumerable<SelectListItem> Colunas { get; set; }
        public IEnumerable<SelectListItem> ColunasSelecionada { get; set; }

    }
}