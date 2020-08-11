using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using Pract1Unidad3.DAL;
using Pract1Unidad3.Models;
using System.Data.SqlClient;
using System.IO;
using System.Xml;
using ClosedXML.Excel;

namespace Practica1Unidad3.Controllers
{
    public class EmpleadosController : Controller
    {
        private EmpleadoContext db = new EmpleadoContext();

        // GET: Empleados
        public ActionResult Index()
        {
            return View(db.Empleados.ToList());
        }

        // GET: Empleados/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Empleado empleado = db.Empleados.Find(id);
            if (empleado == null)
            {
                return HttpNotFound();
            }
            return View(empleado);
        }

        // GET: Empleados/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Empleados/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "EmpleadoID,Nombres,Apellidos,Fecha_Ingreso")] Empleado empleado)
        {
            if (ModelState.IsValid)
            {
                db.Empleados.Add(empleado);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(empleado);
        }

        // GET: Empleados/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Empleado empleado = db.Empleados.Find(id);
            if (empleado == null)
            {
                return HttpNotFound();
            }
            return View(empleado);
        }

        // POST: Empleados/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "EmpleadoID,Nombres,Apellidos,Fecha_Ingreso")] Empleado empleado)
        {
            if (ModelState.IsValid)
            {
                db.Entry(empleado).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(empleado);
        }

        // GET: Empleados/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Empleado empleado = db.Empleados.Find(id);
            if (empleado == null)
            {
                return HttpNotFound();
            }
            return View(empleado);
        }

        // POST: Empleados/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Empleado empleado = db.Empleados.Find(id);
            db.Empleados.Remove(empleado);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        // Agregar nuevo accion para exportar los empleados a un archivo de excel
        public FileResult ExportarEmpleadosCSV() {
            DataTable dt = new DataTable("Primeros3");
            DataTable dt2 = new DataTable("Primeros20");
            DataTable dt3 = new DataTable("NombresJ");
            DataTable dt4 = new DataTable("ApellidosA");

            // Tabla para mostras los primeros 3 empleados seleccionados
            dt.Columns.AddRange(new DataColumn[4] {
                                new DataColumn("Codigo"),
                                new DataColumn("Nombres"),
                                new DataColumn("Apellidos"),
                                new DataColumn("Fecha_Ingreso")
                                });

            var emp = from empleados in db.Empleados.Take(3) select empleados;

            foreach (var e in emp)
            {
                dt.Rows.Add(e.EmpleadoID,e.Nombres,e.Apellidos,e.Fecha_Ingreso);
            }

            // Tabla para mostras los primeros 20 empleados seleccionados
            dt2.Columns.AddRange(new DataColumn[4] {
                                new DataColumn("Codigo"),
                                new DataColumn("Nombres"),
                                new DataColumn("Apellidos"),
                                new DataColumn("Fecha_Ingreso")
                                });

            var primeros = from prim in db.Empleados.Take(20) select prim;

            foreach (var e in primeros)
            {
                dt2.Rows.Add(e.EmpleadoID, e.Nombres, e.Apellidos, e.Fecha_Ingreso);
            }

            // Tabla para mostras los empleados cuyos nombres inicien con J mayuscula
            dt3.Columns.AddRange(new DataColumn[4] {
                                new DataColumn("Codigo"),
                                new DataColumn("Nombres"),
                                new DataColumn("Apellidos"),
                                new DataColumn("Fecha_Ingreso")
                                });

            var nombres = from nomb in db.Empleados
                          where nomb.Nombres.StartsWith("J")
                          select nomb;

            foreach (var n in nombres)
            {
                dt3.Rows.Add(n.EmpleadoID,n.Nombres,n.Apellidos,n.Fecha_Ingreso);
            }

            // Tabla para mostras los empleados cuyos apellidos contengan la letra A mayuscula
            dt4.Columns.AddRange(new DataColumn[4] {
                                new DataColumn("Codigo"),
                                new DataColumn("Nombres"),
                                new DataColumn("Apellidos"),
                                new DataColumn("Fecha_Ingreso")
                                });

            var apellidos = from ap in db.Empleados
                            where ap.Apellidos.Contains("A")
                            select ap;

            foreach (var a in apellidos)
            {
                dt4.Rows.Add(a.EmpleadoID,a.Nombres,a.Apellidos,a.Fecha_Ingreso);
            }

            // Exportar el archivo
            using (XLWorkbook wb = new XLWorkbook()) {
                wb.AddWorksheet(dt);
                wb.AddWorksheet(dt2);
                wb.AddWorksheet(dt3);
                wb.AddWorksheet(dt4);
                using (MemoryStream ms = new MemoryStream() ) {
                    wb.SaveAs(ms);
                    return File(ms.ToArray(),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","Empleados.xlsx");
                };
            };
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
