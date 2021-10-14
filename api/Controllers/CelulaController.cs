using DOC.Data;
using DOC.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mime;
using System.Threading.Tasks;
using System.IO;
using System.Threading;
using System.Text.Json;

namespace DOC.Controllers
{
    [ApiController]
    [Produces(MediaTypeNames.Application.Json)]

    public class CelulaController : ControllerBase
    {
        private readonly DocContext db;

        public CelulaController(DocContext contexto)
        {
            db = contexto;
        }

        [HttpGet("/testecelula")]
        public string TesteCelula()
        {
            return "teste";
        }

        [HttpGet("celula/")]
        public async Task<ActionResult<Celula>> GetCelulas(int tipo = 0)
        {
            try
            {
                if (tipo == 1)
                {

                    return Ok(await db.Celulas.AsNoTracking().ToListAsync());
                }
                else
                {
                    return Ok(await db.Celulas.AsNoTracking().Select(dado => new Celula { Nome = dado.Nome, CelulaId = dado.CelulaId }).ToListAsync());
                }

            }

            catch (Exception ex)
            {
                var errointerno = StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
                return errointerno;
            }

        }

        [HttpGet("celula/{celula}/clientes/dados")]
        public async Task<ActionResult<Cliente>> GetClientesDados(string celula, int tipo = 0)
        {
            try
            {
                if (string.IsNullOrEmpty(celula) == true)
                {
                    return NoContent();
                }
                if (tipo == 0)
                {
                    return Ok(await db.Clientes.AsNoTracking().Join(db.Celulas.AsNoTracking(),
                    dadoscliente => dadoscliente.CelulaId, dadoscelula => dadoscelula.CelulaId, (dadoscliente, dadoscelula) =>
                    new { dadoscliente, dadoscelula }).Where(celulas => celulas.dadoscelula.Nome == celula).
                    Select(clientes => new { clientes.dadoscliente.Nome }).ToListAsync());
                }
                else
                {
                    return Ok(await db.Clientes.AsNoTracking().Join(db.Celulas.AsNoTracking(),
                    dadoscliente => dadoscliente.CelulaId, dadoscelula => dadoscelula.CelulaId, (dadoscliente, dadoscelula) =>
                    new { dadoscliente, dadoscelula }).Join(db.Slas.AsNoTracking(), dadocliente => dadocliente.dadoscliente.ClienteId, dadoSLA => dadoSLA.ClienteId,
                    (cliente, SLA) => new { cliente, SLA }).Where(celulas => celulas.cliente.dadoscelula.Nome == celula).
                    Select(clientes => new { clientes.cliente.dadoscliente.Nome, clientes.SLA.Slaid }).ToListAsync());
                }

            }

            catch (Exception ex)
            {
                var errointerno = StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
                return errointerno;
            }


        }

        [HttpGet("celula/{celula}/clientes")]
        public async Task<ActionResult<Cliente>> GetClientes(string celula)
        {
            try
            {
                return Ok(await db.Clientes.AsNoTracking().Join(db.Celulas.AsNoTracking(), dadoscliente => dadoscliente.CelulaId,
                dadoscelula => dadoscelula.CelulaId, (dadoscliente, dadoscelula) => new { dadoscliente, dadoscelula }).
                Where(celulas => celulas.dadoscelula.Nome == celula).Select(dados => new
                {
                    dados.dadoscliente.ClienteId,
                    dados.dadoscliente.CelulaId,
                    dados.dadoscliente.Nome,
                    dados.dadoscliente.Tipo

                }).ToListAsync());
            }

            catch (Exception ex)
            {
                var errointerno = StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
                return errointerno;
            }

        }

        [HttpGet("celula/{celula}/analistas/nomes")]
        public async Task<ActionResult<Analista>> GetAnalistasNomes(string celula)
        {
            try
            {
                return Ok(await db.Analista.AsNoTracking().Join(db.Celulas.AsNoTracking(),
                dadosanalista => dadosanalista.CelulaId, dadoscelula => dadoscelula.CelulaId,
                (dadosanalista, dadoscelula) => new { dadosanalista, dadoscelula }).
                Where(celulas => celulas.dadoscelula.Nome == celula).Select(analista => new
                {
                    analista.dadosanalista.Nome
                    //inserir ID do usuário.
                }).ToListAsync());

            }

            catch (Exception ex)
            {
                var errointerno = StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
                return errointerno;
            }

        }

        [HttpGet("celula/{celula}/analistas/dados-resumidos")]
        public async Task<ActionResult<Cliente>> GetAnalistasDadosShort(string celula, int tipo, string termo = "")
        {
            try
            {
                #region sp_getcolaboradoresinfo_tipo1
                if (tipo == 1)
                {

                    return Ok(await db.Analista.AsNoTracking().Join(db.Celulas.AsNoTracking(),
                   dadosnalistas => dadosnalistas.CelulaId, dadoscelulas => dadoscelulas.CelulaId,
                   (dadosnalistas, dadoscelulas) => new { dadosnalistas, dadoscelulas }).
                   Where(celulas => celulas.dadoscelulas.Nome == celula).Select(analista => new
                   { analista.dadosnalistas.Nome, analista.dadosnalistas.Email }).ToListAsync());

                }
                #endregion

                #region sp_getcolaboradoresinfo_tipo_Nome
                else if (tipo == 2)
                {
                    return Ok(await db.Analista.AsNoTracking().Join(db.Celulas.AsNoTracking(),
                   dadosanalista => dadosanalista.CelulaId, dadoscelula => dadoscelula.CelulaId,
                   (dadosanalista, dadoscelula) => new { dadosanalista, dadoscelula }).
                   Where(analista => analista.dadosanalista.Nome == termo).
                   Select(consulta => new
                   {
                       NomeAnalista = consulta.dadosanalista.Nome,
                       Email = consulta.dadosanalista.Email,
                       Cargo = consulta.dadosanalista.Cargo,
                       Lideranca = consulta.dadosanalista.Lideranca,
                       NomeCelula = consulta.dadoscelula.Nome

                   }).ToListAsync());
                }
                #endregion

                #region sp_getcolaboradoresinfo_tipo_Email
                else if (tipo == 3)
                {

                    return Ok(await db.Analista.AsNoTracking().Join(db.Celulas.AsNoTracking(),
                    dadosanalista => dadosanalista.CelulaId, dadoscelula => dadoscelula.CelulaId,
                    (dadosanalista, dadoscelula) => new { dadosanalista, dadoscelula }).
                    Where(analista => analista.dadosanalista.Email == termo).
                    Select(consulta => new
                    {
                        NomeAnalista = consulta.dadosanalista.Nome,
                        Email = consulta.dadosanalista.Email,
                        Cargo = consulta.dadosanalista.Cargo,
                        Lideranca = consulta.dadosanalista.Lideranca,
                        NomeCelula = consulta.dadoscelula.Nome

                    }).ToListAsync());

                }

                #endregion

                #region _colaboradorcargocomplexidade

                else if (tipo == 4)
                {

                    return Ok(await db.Analista.AsNoTracking().Join(db.Celulas.AsNoTracking(),
                   dadosnalistas => dadosnalistas.CelulaId, dadoscelulas => dadoscelulas.CelulaId,
                   (dadosnalistas, dadoscelulas) => new { dadosnalistas, dadoscelulas }).
                   Where(celulas => celulas.dadoscelulas.Nome == celula).Select(analista => new
                   { analista.dadosnalistas.Nome, analista.dadosnalistas.CargoComplexidade }).ToListAsync());

                }

                #endregion

                #region sp_getlideremail

                else if (tipo == 5)
                {

                    return Ok(await db.Analista.AsNoTracking().Join(db.Celulas.AsNoTracking(),
                    dadosnalistas => dadosnalistas.CelulaId, dadoscelulas => dadoscelulas.CelulaId,
                    (dadosnalistas, dadoscelulas) => new { dadosnalistas, dadoscelulas }).
                    Where(celulas => celulas.dadoscelulas.Nome == celula && celulas.dadosnalistas.Eliderenca == true)
                    .Select(analista => new { analista.dadosnalistas.Nome, analista.dadosnalistas.Email }).ToListAsync());


                }

                #endregion
                else
                {
                    return NoContent();
                }

            }

            catch (Exception ex)
            {
                var errointerno = StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
                return errointerno;

            }
        }


        [HttpGet("celula/{id}/{celula}/analistas/dados")]
        public async Task<ActionResult<Analista>> GetAnalistasDadosLong(string celula, int id)
        {
            try
            {
                if (id == 0)
                {
                    return BadRequest();
                }

                return Ok(await db.Analista.Where(analista => analista.CelulaId == id).
                Select(dadosanalista => new
                {
                    dadosanalista.AnalistaId,
                    dadosanalista.Nome,
                    dadosanalista.Email,
                    dadosanalista.Cargo,
                    dadosanalista.CargoComplexidade,
                    dadosanalista.Lideranca,
                    dadosanalista.Eliderenca
                }).ToListAsync());



            }

            catch (Exception ex)
            {
                var errointerno = StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
                return errointerno;
            }

        }

        [HttpGet("celula/{celula}/documentos/dados")]
        public async Task<ActionResult<Analista>> GetDocumentosDados(string celula, int tipo, string consulta = "")
        {
            try
            {

                if (tipo == 1)
                {
                    return Ok(await db.Documentos.AsNoTracking().Join(db.Slas.AsNoTracking(),
                   dadosdocumentos => dadosdocumentos.Slaid, SLA => SLA.Slaid,
                   (dadosdocumentos, SLA) => new { documentos = dadosdocumentos, SLAs = SLA }).
                   Join(db.Clientes.AsNoTracking(), sla => sla.SLAs.ClienteId, cliente => cliente.ClienteId,
                   (sla, cliente) => new { slaid = sla, clienteid = cliente }).Join(db.Celulas.AsNoTracking(),
                   dadoscliente => dadoscliente.clienteid.CelulaId, dadoscelula => dadoscelula.CelulaId,
                   (dadoscliente, dadoscelula) => new { cliente = dadoscliente, celula = dadoscelula }).
                   Where(celula_ => celula_.celula.Nome == celula).
                   Select(i => new
                   {

                       documentoNome = i.cliente.slaid.documentos.Nome,
                       i.cliente.slaid.documentos.PrazoMaximoAnalise,
                       i.cliente.slaid.documentos.Tipo,
                       clienteNome = i.cliente.clienteid.Nome,
                       i.cliente.slaid.documentos.Complexidade


                   }).ToListAsync());
                }

                else if (tipo == 2)
                {
                    return Ok(await db.Documentos.AsNoTracking().Join(db.Slas.AsNoTracking(),
                    dadosdocumentos => dadosdocumentos.Slaid, SLA => SLA.Slaid,
                    (dadosdocumentos, SLA) => new { documentos = dadosdocumentos, SLAs = SLA }).
                    Join(db.Clientes.AsNoTracking(), sla => sla.SLAs.ClienteId, cliente => cliente.ClienteId,
                     (sla, cliente) => new { slaid = sla, clientes = cliente }).Join(db.Celulas.AsNoTracking(),
                    dadoscliente => dadoscliente.clientes.CelulaId, dadoscelula => dadoscelula.CelulaId,
                    (dadoscliente, dadoscelula) => new { cliente = dadoscliente, celula = dadoscelula }).
                    Where(celula_ => celula_.celula.Nome == celula && celula_.cliente.clientes.Nome == consulta).
                     Select(i => new
                     {

                         documentoNome = i.cliente.slaid.documentos.Nome,
                         i.cliente.slaid.documentos.PrazoMaximoAnalise,
                         i.cliente.slaid.documentos.Tipo,
                         clienteNome = i.cliente.clientes.Nome,
                         i.cliente.slaid.documentos.Complexidade


                     }).ToListAsync());

                }

                if (tipo == 3)
                {

                    return Ok(await db.Documentos.AsNoTracking().Join(db.Slas.AsNoTracking(),
                   dadosdocumentos => dadosdocumentos.Slaid, SLA => SLA.Slaid,
                   (dadosdocumentos, SLA) => new { documentos = dadosdocumentos, SLAs = SLA }).
                   Join(db.Clientes.AsNoTracking(), sla => sla.SLAs.ClienteId, cliente => cliente.ClienteId,
                   (sla, cliente) => new { slaid = sla, clienteid = cliente }).Join(db.Celulas.AsNoTracking(),
                   dadoscliente => dadoscliente.clienteid.CelulaId, dadoscelula => dadoscelula.CelulaId,
                   (dadoscliente, dadoscelula) => new { cliente = dadoscliente, celula = dadoscelula }).
                   Where(celula_ => celula_.celula.Nome == celula).
                   Select(i => new
                   {
                       i.cliente.slaid.documentos.DocumentoId,
                       documentoNome = i.cliente.slaid.documentos.Nome,
                       i.cliente.slaid.documentos.PrazoMaximoAnalise,
                       i.cliente.slaid.documentos.Tipo,
                       clienteNome = i.cliente.clienteid.Nome,
                       i.cliente.slaid.documentos.Complexidade,
                       clienteTipo = i.cliente.clienteid.Tipo,
                       i.cliente.clienteid.ClienteId,
                       i.cliente.slaid.documentos.TempoMedioAnalise

                   }).ToListAsync());
                }

                else if (tipo == 4)
                {
                    return Ok(await db.Documentos.AsNoTracking().Join(db.Slas.AsNoTracking(),
                    dadosdocumentos => dadosdocumentos.Slaid, SLA => SLA.Slaid,
                    (dadosdocumentos, SLA) => new { documentos = dadosdocumentos, SLAs = SLA }).
                    Join(db.Clientes.AsNoTracking(), sla => sla.SLAs.ClienteId, cliente => cliente.ClienteId,
                    (sla, cliente) => new { slaid = sla, clienteid = cliente }).Join(db.Celulas.AsNoTracking(),
                    dadoscliente => dadoscliente.clienteid.CelulaId, dadoscelula => dadoscelula.CelulaId,
                    (dadoscliente, dadoscelula) => new { cliente = dadoscliente, celula = dadoscelula }).
                    Where(celula_ => celula_.celula.Nome == celula && celula_.cliente.clienteid.Nome == consulta).
                     Select(i => new
                     {

                         documentoNome = i.cliente.slaid.documentos.Nome,
                         i.cliente.slaid.documentos.PrazoMaximoAnalise,
                         i.cliente.slaid.documentos.Tipo,
                         clienteNome = i.cliente.clienteid.Nome,
                         i.cliente.slaid.documentos.Complexidade


                     }).ToListAsync());

                }

                else if (tipo == 5)
                {
                    return Ok(await db.Documentos.AsNoTracking().Join(db.Slas.AsNoTracking(),
                    dadosdocumentos => dadosdocumentos.Slaid, SLA => SLA.Slaid,
                    (dadosdocumentos, SLA) => new { documentos = dadosdocumentos, SLAs = SLA }).
                    Join(db.Clientes.AsNoTracking(), sla => sla.SLAs.ClienteId, cliente => cliente.ClienteId,
                    (sla, cliente) => new { slaid = sla, clienteid = cliente }).Join(db.Celulas.AsNoTracking(),
                    dadoscliente => dadoscliente.clienteid.CelulaId, dadoscelula => dadoscelula.CelulaId,
                    (dadoscliente, dadoscelula) => new { cliente = dadoscliente, celula = dadoscelula }).
                    Where(celula_ => celula_.celula.Nome == celula && celula_.cliente.clienteid.Nome == consulta).
                     Select(i => new
                     {

                         documentoNome = i.cliente.slaid.documentos.Nome,
                         i.cliente.slaid.documentos.PrazoMaximoAnalise,
                         i.cliente.slaid.documentos.Tipo,
                         clienteNome = i.cliente.clienteid.Nome,
                         i.cliente.slaid.documentos.Complexidade


                     }).ToListAsync());

                }

                else
                {
                    return Ok(await db.Documentos.AsNoTracking().ToListAsync());
                }


            }

            catch (Exception ex)
            {
                var errointerno = StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
                return errointerno;
            }

        }

        [HttpGet("celula/{celulaid}/{celula}/documentos/prazo")]
        public async Task<ActionResult<Documento>> GetDocumentosPrazo(string celula, int celulaid = 0)
        {
            try
            {
                if (celulaid > 0)
                {
                    return Ok(await db.Documentos.AsNoTracking().Join(db.Slas, dadosdocumentos => dadosdocumentos.Slaid,
                    dadossla => dadossla.Slaid, (dadosdocumentos, dadossla) => new { documentos = dadosdocumentos, SLAs = dadossla }).
                    Join(db.Clientes, dadosSLA => dadosSLA.SLAs.ClienteId, dadoscliente => dadoscliente.ClienteId, (dadosSLA, dadoscliente) =>
                    new { clientes = dadoscliente, SLA = dadosSLA }).Join(db.Celulas, dadosclientes => dadosclientes.clientes.CelulaId,
                    dadoscelulas => dadoscelulas.CelulaId, (dadosclientes, dadoscelulas) => new { dadocliente = dadosclientes, dadocelula = dadoscelulas })
                    .Where(termopesq => termopesq.dadocelula.CelulaId == celulaid && new[] { 0, 1 }.Contains(termopesq.dadocliente.SLA.documentos.Tipo))
                    .Select(consulta => new { consulta.dadocliente.SLA.documentos.PrazoMaximoAnalise }).Distinct().ToListAsync());
                }

                else
                {
                    return Ok(await db.Documentos.AsNoTracking().Join(db.Slas, dadosdocumentos => dadosdocumentos.Slaid,
                    dadossla => dadossla.Slaid, (dadosdocumentos, dadossla) => new { documentos = dadosdocumentos, SLAs = dadossla }).
                    Join(db.Clientes, dadosSLA => dadosSLA.SLAs.ClienteId, dadoscliente => dadoscliente.ClienteId, (dadosSLA, dadoscliente) =>
                    new { clientes = dadoscliente, SLA = dadosSLA }).Join(db.Celulas, dadosclientes => dadosclientes.clientes.CelulaId,
                    dadoscelulas => dadoscelulas.CelulaId, (dadosclientes, dadoscelulas) => new { dadocliente = dadosclientes, dadocelula = dadoscelulas })
                    .Where(termopesq => termopesq.dadocelula.Nome == celula && new[] { 0, 1 }.Contains(termopesq.dadocliente.SLA.documentos.Tipo))
                    .Select(consulta => new { consulta.dadocliente.SLA.documentos.PrazoMaximoAnalise }).Distinct().ToListAsync());
                }

            }

            catch (Exception ex)
            {
                var errointerno = StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
                return errointerno;
            }

        }

        [HttpPost("celula/novo")]
        public async Task<ActionResult<Celula>> PostCliente([FromBody] Celula celula, int id)
        {
            try
            {

                var novocelula = db.Celulas.AddAsync(celula);
                return Ok(await db.SaveChangesAsync());
            }

            catch (Exception ex)
            {
                var errointerno = StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
                return errointerno;
            }

        }


        [HttpPost("celula/{celula}/cliente/novo")]
        public async Task<ActionResult<Cliente>> PostCliente([FromBody] Cliente cliente, string celula)
        {
            try
            {
                if (cliente.CelulaId == 0)
                {
                    return BadRequest();
                }

                cliente.Nome = System.Web.HttpUtility.UrlDecode(cliente.Nome);
                var novocliente = await db.Clientes.AddAsync(cliente);
                await db.SaveChangesAsync();
                Sla SLA = new Sla();
                SLA.ClienteId = cliente.ClienteId;
                await db.Slas.AddAsync(SLA);
                await db.SaveChangesAsync();
                return Created(nameof(GetClientes), new { id = cliente.ClienteId });
            }

            catch (Exception ex)
            {
                var errointerno = StatusCode(StatusCodes.Status500InternalServerError, new { erro = ex.Message, local = nameof(PostCliente) });
                return errointerno;
            }

        }

        [HttpPost("celula/{id}/{celula}/analista/novo")]
        public async Task<ActionResult<Analista>> PostAnalista([FromBody] Analista analista, int id)
        {
            try
            {
                analista.Nome = System.Web.HttpUtility.UrlDecode(analista.Nome);
                analista.Email = System.Web.HttpUtility.UrlDecode(analista.Email);
                analista.Lideranca = System.Web.HttpUtility.UrlDecode(analista.Lideranca);
                await db.Analista.AddAsync(analista);
                await db.SaveChangesAsync();
                return Ok(new { analistaId = analista.AnalistaId });
            }

            catch (Exception ex)
            {
                var errointerno = StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
                return errointerno;
            }
        }

        [HttpPost("celula/{id}/{celula}/documento/novo")]
        public async Task<ActionResult<Documento>> PostDocumento([FromBody] Documento documento, int id, int ambiente = 0)
        {
            try
            {
                if (documento.Slaid == 0)
                {

                    var SLAID = await db.Slas.AsNoTracking().Join(db.Clientes.AsNoTracking(),
                    dadosSLA => dadosSLA.ClienteId, dadosCliente => dadosCliente.ClienteId, (dadosSla, dadoscliente) =>
                     new { dadosSla, dadoscliente }).Where(cliente => cliente.dadoscliente.Nome == documento.Cliente
                     && cliente.dadoscliente.Tipo == ambiente).Select(SLA => new { SLA.dadosSla.Slaid }).FirstOrDefaultAsync();
                    documento.Slaid = Convert.ToInt32(SLAID.Slaid);
                }
                string data = documento.TempoMedioAnaliseBruto.ToString();
                data = System.Web.HttpUtility.UrlDecode(data);
                documento.TempoMedioAnaliseBruto = Convert.ToDateTime(data);
                documento.TempoMedioAnalise = documento.TempoMedioAnaliseBruto.TimeOfDay;
                await db.Documentos.AddAsync(documento);
                await db.SaveChangesAsync();
                return Ok(new { tipo = 1, documentoNovo = documento.DocumentoId });

            }

            catch (Exception ex)
            {
                var errointerno = StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
                return errointerno;
            }
        }

        [HttpPatch("celula/{id}/{celula}/documento/alterar/{documentoid}")]
        public async Task<ActionResult<Documento>> PatchDocumento([FromBody] Documento documento, string celula, int id, int documentoid, int ambiente = 0)
        {

            if (ModelState.IsValid)
            {
                documento.Slaid = await db.Documentos.Where(dado => dado.DocumentoId == documentoid).Select(I => I.Slaid).SingleOrDefaultAsync();
                db.Entry(documento).State = EntityState.Modified;

                try
                {
                    return Ok(new { documentosAlterados = await db.SaveChangesAsync() });

                }

                catch (DbUpdateConcurrencyException dbex)
                {
                    return NotFound(dbex);
                }
            }
            else
            {
                return NotFound();
            }

        }

        [HttpPatch("celula/{id}/{celula}/analista/alterar/{analistaid}")]
        public async Task<ActionResult<Analista>> PatchAnalista([FromBody] Analista analista, string celula, int id, int analistaid, int ambiente = 0)
        {

            if (ModelState.IsValid)
            {
                analista.Nome = System.Web.HttpUtility.UrlDecode(analista.Nome);
                analista.Email = System.Web.HttpUtility.UrlDecode(analista.Email);
                analista.Lideranca = System.Web.HttpUtility.UrlDecode(analista.Lideranca);
                if (analista.Eliderenca == true)
                {
                    analista.CargoComplexidade = 5;
                }

                if (analista.Eliderenca == false)
                {
                    string[] lideres = await db.Analista.AsNoTracking().Where(analistas => analistas.Eliderenca == true &&
                    analistas.CelulaId == id).Select(analistas => analistas.Nome).ToArrayAsync();
                    int qtdLideres = lideres.Length;
                    analista.CargoComplexidade = 5;


                    for (int i = 0; i < qtdLideres; i++)
                    {
                        if (i == qtdLideres - 1)
                        {
                            analista.Lideranca += lideres[i];
                        }

                        else if (i != qtdLideres)
                        {
                            analista.Lideranca += lideres[i] + ",";
                        }

                    }
                }



                db.Entry(analista).State = EntityState.Modified;

                try
                {
                    await db.SaveChangesAsync();

                    if (analista.Eliderenca == true)
                    {
                        var lideres_ = await db.Analista.AsNoTracking().Where(analistas => analistas.Eliderenca == true &&
                        analistas.CelulaId == id).Select(analistas => analistas.Nome).ToArrayAsync();
                        int qtdLideres_ = lideres_.Length;



                        for (int i = 0; i < qtdLideres_; i++)
                        {
                            if (i == qtdLideres_ - 1)
                            {
                                analista.Lideranca += lideres_[i];
                            }

                            else if (i != qtdLideres_)
                            {
                                analista.Lideranca += lideres_[i] + ",";
                            }

                        }
                        string SQLUpdate = "update [dbo].[Analista] set [Lideranca] =@Lideranca where [CelulaID] = @CelulaID";

                        var atualizacao = db.Database.ExecuteSqlRawAsync(SQLUpdate,
                            new Microsoft.Data.SqlClient.SqlParameter("@Lideranca", analista.Lideranca),
                            new Microsoft.Data.SqlClient.SqlParameter("@CelulaID", analista.CelulaId));

                        if (atualizacao.Result > 0)
                        {
                            return Ok(new { registroAlterado = 1 });
                        }
                    }
                    else
                    {
                        return Ok(new { registroAlterado = 1 });
                    }

                    return BadRequest();


                }

                catch (Exception ex)
                {
                    return NotFound(ex);
                }
            }
            else
            {
                return NotFound();
            }

        }

        [HttpPatch("celula/{id}/{celula}/cliente/alterar/{clienteid}")]
        public async Task<ActionResult<Cliente>> PatchCliente([FromBody] Cliente cliente, string celula, int id, int clienteid, int ambiente = 0)
        {

            cliente.Nome = System.Web.HttpUtility.UrlDecode(cliente.Nome);
            db.Entry(cliente).State = EntityState.Modified;

            try
            {

                return Ok(new { registroAlterado = await db.SaveChangesAsync() });

            }

            catch (DbUpdateConcurrencyException dbex)
            {
                return NotFound(dbex);
            }


        }

        [HttpPost("celula/{celulaid}/{celula}/documento/alterar/")]
        public async Task<ActionResult<Documento>> EditDocumento([FromBody] Documento documento, string celula)
        {

            try
            {

                celula = System.Web.HttpUtility.UrlDecode(celula);

                var documentoDados = await db.Documentos.AsNoTracking().Where(doc => doc.Nome == documento.Nome).FirstOrDefaultAsync();

                if (documentoDados != null)
                {
                    var SLAID = await db.Slas.AsNoTracking().Join(db.Clientes.AsNoTracking(), dadosSLA => dadosSLA.ClienteId,
                    dadoscliente => dadoscliente.ClienteId, (dadosSLA, dadoscliente) =>
                    new { SLA = dadosSLA, cliente = dadoscliente }).Join(db.Celulas.AsNoTracking(),
                    SLA => SLA.cliente.CelulaId, celulas => celulas.CelulaId, (SLA, celula) =>
                    new { SLA, Celula = celula }).Where(dados => dados.SLA.cliente.Nome == documento.Cliente && dados.Celula.Nome == celula).
                    Select(SLA => SLA.SLA.SLA.Slaid).FirstOrDefaultAsync();
                    documento.PrazoMaximoAnalise = documentoDados.PrazoMaximoAnalise;
                    documento.Slaid = Convert.ToInt32(SLAID);
                    documento.Complexidade = documentoDados.Complexidade;
                    documento.Tipo = documentoDados.Tipo;
                    documento.TempoMedioAnalise = documentoDados.TempoMedioAnalise;
                    await db.AddAsync(documento);
                    await db.SaveChangesAsync();
                    return Ok(new { tipo = 1, documentoNovo = documento.DocumentoId });
                }

                else
                {
                    return Ok(new { tipo = 2 });
                }

            }

            catch (Exception ex)
            {
                var errointerno = StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
                return errointerno;
            }
        }
        [HttpPost]
        public async Task<ActionResult> DistribuirDocumentosTipo1([FromQuery] IFormFile arquivoJSON, CancellationToken token)
        {

            try
            {
                string extensao = Path.GetExtension(arquivoJSON.FileName);
                if (extensao != ".json")
                {
                    return BadRequest("Formato não suportado");
                }

                string json = null;

                using (StreamReader sr = new(arquivoJSON.FileName)) 
                {
                    json = await sr.ReadToEndAsync();
                }   

                List<DistribDocumento> distribuicao = JsonSerializer.Deserialize<List<DistribDocumento>>(json);
                List<DistribDocumento> documentosJaAtribuidosAnteriormente = new();
                List<DistribDocumento> documentosDistribuidos = new();

                foreach (var documento in distribuicao)
                {
                    var analista = await db.DocumentoAvalidars.FirstOrDefaultAsync(x => x.DocumentoId == documento.Protocolo, token);
                    if (analista.AnalistaId != 0)
                    {
                        db.Entry(documento).State = EntityState.Modified;

                        documentosDistribuidos.Add(new DistribDocumento
                        {
                            Analista = analista.AnalistaId.ToString(),
                            Protocolo = documento.Protocolo
                        });

                    }

                    else
                    {  
                        documentosJaAtribuidosAnteriormente.Add(new DistribDocumento
                        {
                            Analista = analista.AnalistaId.ToString(),
                            Protocolo = documento.Protocolo
                        });
                    }
                }

                return Ok(new { atribuidos = documentosDistribuidos, atribudosAnteriormente = documentosJaAtribuidosAnteriormente });
            }

            catch(Exception ex)
            {
                var errointerno = StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
                return errointerno;
            }
           

        }


    }

}





