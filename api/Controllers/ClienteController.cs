using DOC.Data;
using DOC.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DOC.Controllers
{
    [ApiController]
    [Produces("application/json")]
    public class ClienteController : ControllerBase
    {
        private readonly DocContext db;
      
        public ClienteController(DocContext contexto)
        {
            db = contexto;
        }

        [HttpGet("cliente/{celula}")]
        public async Task<ActionResult<Cliente>> GetClientesNome(string celula)
        {
            try
            {
                if (string.IsNullOrEmpty(celula) == true)
                {
                    return NoContent();
                }

                return Ok(await db.Clientes.AsNoTracking().Join(db.Celulas.AsNoTracking(),
                dadoscliente => dadoscliente.CelulaId, dadoscelula => dadoscelula.CelulaId, (dadoscliente, dadoscelula) =>
                new { dadoscliente, dadoscelula }).Where(celulas => celulas.dadoscelula.Nome == celula).
                Select(clientes => new { clientes.dadoscliente.Nome }).ToListAsync());

            }

            catch (Exception ex)
            {
                var errointerno = StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
                return errointerno;
            }


        }
        
        [HttpGet("cliente/{celula}/clientes")]
        public async Task<ActionResult<Cliente>> GetClientes(string celula)
        {
            try
            {
                return Ok(await db.Clientes.AsNoTracking().Join(db.Celulas, dadoscliente => dadoscliente.CelulaId,
                dadoscelula => dadoscelula.CelulaId, (dadoscliente, dadoscelula) => new { dadoscliente, dadoscelula }).
                Where(celulas => celulas.dadoscelula.Nome == celula).Select(dados => new
                {
                    dados.dadoscliente.ClienteId,
                    dados.dadoscliente.CelulaId,
                    dados.dadoscliente.Nome,
                    dados.dadoscliente.Tipo
                }).ToListAsync());
            }

            catch(Exception ex)
            {
                var errointerno = StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
                return errointerno;
            }

        }

        [HttpPost("cliente/{celula}/{id}/new")]
        public async Task<ActionResult<Cliente>> PostCliente([FromBody] Cliente cliente, int id)
        {
            try
            {
                cliente.CelulaId = id;
                await db.Clientes.AddAsync(cliente);
                await db.SaveChangesAsync();
                return Ok();
            }

            catch(Exception ex)
            {
                 var errointerno = StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
                return errointerno;
            }
            

            
        }   

    }


    
}
