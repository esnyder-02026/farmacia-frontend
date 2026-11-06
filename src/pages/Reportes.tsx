import React, { useState, useRef } from 'react';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import ExcelJS from 'exceljs';

import {
  getCategoriasActivas,
  getMedicamentosFiltro,
  getReporteMedicamentoCategoria,
  getMedicamentosStockBajo,
  getReporteClienteVenta,
  getReporteClienteVentaDetalle,
} from '../services/reporteServices';

export const Reportes: React.FC = () => {
  const [mensaje, setMensaje] = useState<string>('');
  const [tipoMensaje, setTipoMensaje] = useState<'exito' | 'error'>('exito');

  // Estados para modales de filtro
  const [filtroNombre, setFiltroNombre] = useState<string>('crema');
  const [filtroStock, setFiltroStock] = useState<number>(10);

  const dialogNombreRef = useRef<HTMLDialogElement>(null);
  const dialogStockRef = useRef<HTMLDialogElement>(null);

  const mostrarNotificacion = (text: string, esError = false) => {
    setMensaje(text);
    setTipoMensaje(esError ? 'error' : 'exito');
  };

  // --- HELPER PDF BASE ---
  const generarPDFBase = (
    titulo: string, 
    columnas: string[], 
    filas: (string | number)[][], 
    nombreArchivo: string,
    totalGeneral?: number
  ) => {
    const doc = new jsPDF({ orientation: 'landscape', format: 'a4' });

    doc.setFillColor(13, 148, 136); // Teal primary
    doc.rect(0, 0, 297, 18, 'F');
    
    doc.setFontSize(14);
    doc.setTextColor(255, 255, 255);
    doc.setFont('helvetica', 'bold');
    doc.text('FARMACIA SANTA CRUZ • SISTEMA DE REPORTES', 14, 12);

    doc.setFontSize(16);
    doc.setTextColor(13, 148, 136);
    doc.text(titulo, 14, 28);

    doc.setFontSize(9);
    doc.setTextColor(100);
    doc.setFont('helvetica', 'normal');
    doc.text(`Fecha de generación: ${new Date().toLocaleDateString()} ${new Date().toLocaleTimeString()}`, 14, 34);

    autoTable(doc, {
      head: [columnas],
      body: filas,
      foot: totalGeneral !== undefined ? [["", "", "", "", "Total general", `Q ${totalGeneral.toFixed(2)}`]] : undefined,
      footStyles: { fillColor: [13, 148, 136], textColor: 255, fontStyle: 'bold' },
      startY: 38,
      theme: 'grid',
      headStyles: { fillColor: [13, 148, 136], textColor: 255, fontStyle: 'bold' },
      alternateRowStyles: { fillColor: [240, 253, 250] },
      styles: { fontSize: 9, cellPadding: 3 },
    });

    const totalPaginas = (doc as any).internal.getNumberOfPages();
    const ancho = doc.internal.pageSize.getWidth();
    const alto = doc.internal.pageSize.getHeight();

    for (let i = 1; i <= totalPaginas; i++) {
      doc.setPage(i);
      doc.setFontSize(8);
      doc.setTextColor(150);
      doc.text(`Página ${i} de ${totalPaginas}`, ancho / 2, alto - 8, { align: 'center' });
    }

    doc.save(`${nombreArchivo}.pdf`);
  };

  // --- HELPER EXCEL BASE ---
  const generarExcelBase = async (
    titulo: string,
    columnas: string[],
    filas: (string | number)[][],
    nombreArchivo: string,
    formatoPreciosCols: number[] = [],
    evaluarStockRojo: boolean = false,
    formulaSumaCol?: number
  ) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Reporte Farmacia');

    const filaTitulo = worksheet.addRow([titulo]);
    filaTitulo.font = { bold: true, size: 16, color: { argb: 'FF0D9488' } };

    worksheet.addRow([`Fecha de generación: ${new Date().toLocaleDateString()}`]);
    worksheet.addRow([]);

    const filaEncabezados = worksheet.addRow(columnas);
    filaEncabezados.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0D9488' } };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    filas.forEach((fila, index) => {
      const row = worksheet.addRow(fila);

      if (index % 2 === 1) {
        row.eachCell((cell) => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0FDF4' } };
        });
      }

      if (evaluarStockRojo) {
        const stockVal = Number(fila[2]);
        if (stockVal < 5) {
          row.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
          row.getCell(3).font = { color: { argb: 'FF991B1B' }, bold: true };
        }
      }

      formatoPreciosCols.forEach((colIdx) => {
        row.getCell(colIdx).numFmt = '"Q"#,##0.00';
      });

      row.eachCell((cell) => {
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        };
      });
    });

    if (formulaSumaCol) {
      const ultima = worksheet.rowCount;
      const colLetra = String.fromCharCode(64 + formulaSumaCol);
      const filaTotalArr = Array(columnas.length).fill('');
      filaTotalArr[formulaSumaCol - 2] = 'Total general';
      filaTotalArr[formulaSumaCol - 1] = { formula: `SUM(${colLetra}5:${colLetra}${ultima})` };

      const filaTotal = worksheet.addRow(filaTotalArr);
      filaTotal.font = { bold: true };
      worksheet.getColumn(formulaSumaCol).numFmt = '"Q"#,##0.00';
    }

    worksheet.columns.forEach((col) => { col.width = 24; });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${nombreArchivo}.xlsx`;
    a.click();
    window.URL.revokeObjectURL(url);
  };

  // --- HANDLERS ---
  const rep1PDF = async () => {
    try {
      const res = await getCategoriasActivas();
      const filas = res.data.map((c) => [c.nombre, c.descripcion || '-']);
      generarPDFBase('Reporte de Categorías Activas', ['Nombre', 'Descripción'], filas, 'categoriasActivas');
      mostrarNotificacion('Categorías Activas descargado en PDF.');
    } catch {
      mostrarNotificacion('Error al conectar con la API de categorías.', true);
    }
  };

  const rep1Excel = async () => {
    try {
      const res = await getCategoriasActivas();
      const filas = res.data.map((c) => [c.nombre, c.descripcion || '-']);
      await generarExcelBase('Reporte de Categorías Activas', ['Nombre', 'Descripción'], filas, 'categoriasActivas');
      mostrarNotificacion('Categorías Activas descargado en Excel.');
    } catch {
      mostrarNotificacion('Error al conectar con la API de categorías.', true);
    }
  };

  const rep2PDF = async () => {
    try {
      const res = await getMedicamentosFiltro(filtroNombre);
      const lista = Array.isArray(res.data) ? res.data : [];
      if (lista.length === 0) {
        mostrarNotificacion(`No se encontraron medicamentos con "${filtroNombre}".`, true);
        return;
      }
      const filas = lista.map((m) => [m.nombre, m.laboratorio, Number(m.precio || 0).toFixed(2), m.stock]);
      generarPDFBase(`Medicamentos con filtro: "${filtroNombre}"`, ['Nombre', 'Laboratorio', 'Precio (Q)', 'Stock'], filas, 'medicamentosPorNombre');
      dialogNombreRef.current?.close();
      mostrarNotificacion('Medicamentos por Nombre descargado en PDF.');
    } catch {
      mostrarNotificacion('Error al obtener medicamentos.', true);
    }
  };

  const rep2Excel = async () => {
    try {
      const res = await getMedicamentosFiltro(filtroNombre);
      const lista = Array.isArray(res.data) ? res.data : [];
      if (lista.length === 0) {
        mostrarNotificacion(`No se encontraron medicamentos con "${filtroNombre}".`, true);
        return;
      }
      const filas = lista.map((m) => [m.nombre, m.laboratorio, m.precio, m.stock]);
      await generarExcelBase(`Medicamentos con filtro: "${filtroNombre}"`, ['Nombre', 'Laboratorio', 'Precio (Q)', 'Stock'], filas, 'medicamentosPorNombre', [3]);
      dialogNombreRef.current?.close();
      mostrarNotificacion('Medicamentos por Nombre descargado en Excel.');
    } catch {
      mostrarNotificacion('Error al obtener medicamentos.', true);
    }
  };

  const rep3PDF = async () => {
    try {
      const res = await getReporteMedicamentoCategoria();
      const filas = res.data.map((mc) => [mc.nombreMedicamento, mc.laboratorio, mc.nombreCategoria, mc.precio.toFixed(2), mc.stock]);
      generarPDFBase('Reporte Medicamento - Categoría', ['Medicamento', 'Laboratorio', 'Categoría', 'Precio (Q)', 'Stock'], filas, 'medicamentoCategoria');
      mostrarNotificacion('Medicamento + Categoría descargado en PDF.');
    } catch {
      mostrarNotificacion('Error al cargar reporte 3.', true);
    }
  };

  const rep3Excel = async () => {
    try {
      const res = await getReporteMedicamentoCategoria();
      const filas = res.data.map((mc) => [mc.nombreMedicamento, mc.laboratorio, mc.nombreCategoria, mc.precio, mc.stock]);
      await generarExcelBase('Reporte Medicamento - Categoría', ['Medicamento', 'Laboratorio', 'Categoría', 'Precio (Q)', 'Stock'], filas, 'medicamentoCategoria', [4]);
      mostrarNotificacion('Medicamento + Categoría descargado en Excel.');
    } catch {
      mostrarNotificacion('Error al cargar reporte 3.', true);
    }
  };

  const rep4PDF = async () => {
    try {
      const res = await getMedicamentosStockBajo(filtroStock);
      const lista = Array.isArray(res.data) ? res.data : [];
      if (lista.length === 0) {
        mostrarNotificacion(`No hay medicamentos con stock menor a ${filtroStock}.`, true);
        return;
      }
      const filas = lista.map((m) => [m.nombre, m.laboratorio, m.stock]);
      generarPDFBase(`Reporte Stock Bajo (Menor a ${filtroStock})`, ['Medicamento', 'Laboratorio', 'Stock'], filas, 'stockBajo');
      dialogStockRef.current?.close();
      mostrarNotificacion('Medicamentos Stock Bajo descargado en PDF.');
    } catch {
      mostrarNotificacion('Error al obtener stock bajo.', true);
    }
  };

  const rep4Excel = async () => {
    try {
      const res = await getMedicamentosStockBajo(filtroStock);
      const lista = Array.isArray(res.data) ? res.data : [];
      if (lista.length === 0) {
        mostrarNotificacion(`No hay medicamentos con stock menor a ${filtroStock}.`, true);
        return;
      }
      const filas = lista.map((m) => [m.nombre, m.laboratorio, m.stock]);
      await generarExcelBase(`Reporte Stock Bajo (Menor a ${filtroStock})`, ['Medicamento', 'Laboratorio', 'Stock'], filas, 'stockBajo', [], true);
      dialogStockRef.current?.close();
      mostrarNotificacion('Medicamentos Stock Bajo descargado en Excel.');
    } catch {
      mostrarNotificacion('Error al obtener stock bajo.', true);
    }
  };

  const rep5PDF = async () => {
    try {
      const res = await getReporteClienteVenta();
      const filas = res.data.map((v) => [`${v.nombreCliente} ${v.apellidoCliente}`, v.fechaVenta, v.total.toFixed(2)]);
      generarPDFBase('Reporte Ventas por Cliente', ['Cliente', 'Fecha Venta', 'Total (Q)'], filas, 'ventasPorCliente');
      mostrarNotificacion('Ventas por Cliente descargado en PDF.');
    } catch {
      mostrarNotificacion('Error al obtener reporte 5.', true);
    }
  };

  const rep5Excel = async () => {
    try {
      const res = await getReporteClienteVenta();
      const filas = res.data.map((v) => [`${v.nombreCliente} ${v.apellidoCliente}`, v.fechaVenta, v.total]);
      await generarExcelBase('Reporte Ventas por Cliente', ['Cliente', 'Fecha Venta', 'Total (Q)'], filas, 'ventasPorCliente', [3]);
      mostrarNotificacion('Ventas por Cliente descargado en Excel.');
    } catch {
      mostrarNotificacion('Error al obtener reporte 5.', true);
    }
  };

  const rep6PDF = async () => {
    try {
      const res = await getReporteClienteVentaDetalle();
      const totalGen = res.data.reduce((sum, d) => sum + d.subtotal, 0);
      const filas = res.data.map((vd) => [
        `${vd.nombreCliente} ${vd.apellidoCliente}`,
        vd.fechaVenta,
        vd.nombreMedicamento,
        vd.cantidad,
        vd.precioUnitario.toFixed(2),
        vd.subtotal.toFixed(2),
      ]);
      generarPDFBase('Reporte Detalle de Ventas por Cliente', ['Cliente', 'Fecha', 'Medicamento', 'Cantidad', 'Precio Unit.', 'SubTotal (Q)'], filas, 'ventasClienteDetalle', totalGen);
      mostrarNotificacion('Ventas Cliente Detalle descargado en PDF.');
    } catch {
      mostrarNotificacion('Error al obtener reporte 6.', true);
    }
  };

  const rep6Excel = async () => {
    try {
      const res = await getReporteClienteVentaDetalle();
      const filas = res.data.map((vd) => [
        `${vd.nombreCliente} ${vd.apellidoCliente}`,
        vd.fechaVenta,
        vd.nombreMedicamento,
        vd.cantidad,
        vd.precioUnitario,
        vd.subtotal,
      ]);
      await generarExcelBase('Reporte Detalle de Ventas por Cliente', ['Cliente', 'Fecha', 'Medicamento', 'Cantidad', 'Precio Unit.', 'SubTotal (Q)'], filas, 'ventasClienteDetalle', [5, 6], false, 6);
      mostrarNotificacion('Ventas Cliente Detalle descargado en Excel.');
    } catch {
      mostrarNotificacion('Error al obtener reporte 6.', true);
    }
  };

  return (
    <div className="min-h-screen bg-slate-100 text-slate-800 font-sans pb-16">
      {/* Header con degradado farmacéutico */}
      <header className="bg-gradient-to-r from-teal-800 via-teal-700 to-emerald-800 text-white shadow-md">
        <div className="max-w-6xl mx-auto px-6 py-6 flex flex-col md:flex-row items-center justify-between gap-4">
          <div className="flex items-center space-x-4">
            <div className="bg-white/10 backdrop-blur-md p-3 rounded-2xl border border-white/20 text-2xl font-bold">
              ⚕️
            </div>
            <div>
              <h1 className="text-2xl font-extrabold tracking-wide">FARMACIA SANTA CRUZ</h1>
              <p className="text-xs text-teal-100 font-medium">Módulo Oficial de Reportes e Impresión • Spring Boot & React</p>
            </div>
          </div>
          <div className="bg-emerald-500/20 backdrop-blur-md px-4 py-2 rounded-xl border border-emerald-400/30 flex items-center gap-2">
            <span className="w-2.5 h-2.5 bg-emerald-400 rounded-full animate-pulse"></span>
            <span className="text-xs font-bold text-emerald-100">API Conectada</span>
          </div>
        </div>
      </header>

      {/* Contenido Principal */}
      <main className="max-w-6xl mx-auto px-6 mt-8">
        
        {/* Banner Informativo */}
        <div className="bg-white rounded-2xl p-6 shadow-sm border border-slate-200/80 mb-8 flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <h2 className="text-xl font-bold text-slate-900">Centro de Generación Documental</h2>
            <p className="text-xs text-slate-500 mt-1">
              Exportación de informes oficiales formateados en tiempo real a PDF (A4 Horizontal) y hojas de cálculo Excel.
            </p>
          </div>
          <span className="text-xs font-semibold bg-teal-50 text-teal-700 px-3 py-1.5 rounded-lg border border-teal-200 w-fit">
            6 Reportes Disponibles
          </span>
        </div>

        {/* Notificación Flotante */}
        {mensaje && (
          <div
            className={`p-4 rounded-xl mb-6 text-xs font-bold transition-all shadow-sm flex items-center justify-between ${
              tipoMensaje === 'error'
                ? 'bg-rose-50 border border-rose-200 text-rose-800'
                : 'bg-emerald-50 border border-emerald-200 text-emerald-800'
            }`}
          >
            <div className="flex items-center gap-2">
              <span>{tipoMensaje === 'error' ? '❌' : '✅'}</span>
              <span>{mensaje}</span>
            </div>
            <button
              onClick={() => setMensaje('')}
              className="text-xs text-slate-500 hover:text-slate-800 font-bold ml-4"
            >
              ✕
            </button>
          </div>
        )}

        {/* Grid de Tarjetas */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">

          {/* Tarjeta 1 */}
          <div className="bg-white rounded-2xl p-6 border border-slate-200/80 shadow-sm hover:shadow-md transition-all flex flex-col justify-between group">
            <div>
              <div className="flex items-center justify-between mb-3">
                <span className="bg-teal-50 text-teal-700 border border-teal-200 text-[10px] font-black tracking-wider uppercase px-2.5 py-1 rounded-md">
                  REPORTE 01
                </span>
                <span className="text-2xl group-hover:scale-110 transition-transform">🏷️</span>
              </div>
              <h3 className="font-bold text-slate-900 text-base mb-1">Categorías Activas</h3>
              <p className="text-xs text-slate-500 leading-relaxed mb-6">Listado oficial de familias de medicamentos registradas en la base de datos.</p>
            </div>
            <div className="grid grid-cols-2 gap-2 pt-3 border-t border-slate-100">
              <button
                onClick={rep1PDF}
                className="bg-gradient-to-r from-rose-500 to-rose-600 hover:from-rose-600 hover:to-rose-700 text-white font-bold text-xs py-2.5 rounded-xl shadow-sm transition"
              >
                📄 PDF
              </button>
              <button
                onClick={rep1Excel}
                className="bg-gradient-to-r from-emerald-600 to-teal-700 hover:from-emerald-700 hover:to-teal-800 text-white font-bold text-xs py-2.5 rounded-xl shadow-sm transition"
              >
                📊 Excel
              </button>
            </div>
          </div>

          {/* Tarjeta 2 */}
          <div className="bg-white rounded-2xl p-6 border border-slate-200/80 shadow-sm hover:shadow-md transition-all flex flex-col justify-between group">
            <div>
              <div className="flex items-center justify-between mb-3">
                <span className="bg-teal-50 text-teal-700 border border-teal-200 text-[10px] font-black tracking-wider uppercase px-2.5 py-1 rounded-md">
                  REPORTE 02
                </span>
                <span className="text-2xl group-hover:scale-110 transition-transform">🔍</span>
              </div>
              <h3 className="font-bold text-slate-900 text-base mb-1">Medicamentos por Nombre</h3>
              <p className="text-xs text-slate-500 leading-relaxed mb-6">Filtro interactivo por coincidencia de texto e impresión personalizada.</p>
            </div>
            <div className="pt-3 border-t border-slate-100">
              <button
                onClick={() => dialogNombreRef.current?.showModal()}
                className="w-full bg-slate-900 hover:bg-slate-800 text-white font-bold text-xs py-2.5 rounded-xl shadow-sm transition flex items-center justify-center gap-2"
              >
                ⚙️ Configurar Filtro
              </button>
            </div>
          </div>

          {/* Tarjeta 3 */}
          <div className="bg-white rounded-2xl p-6 border border-slate-200/80 shadow-sm hover:shadow-md transition-all flex flex-col justify-between group">
            <div>
              <div className="flex items-center justify-between mb-3">
                <span className="bg-teal-50 text-teal-700 border border-teal-200 text-[10px] font-black tracking-wider uppercase px-2.5 py-1 rounded-md">
                  REPORTE 03
                </span>
                <span className="text-2xl group-hover:scale-110 transition-transform">💊</span>
              </div>
              <h3 className="font-bold text-slate-900 text-base mb-1">Medicamento + Categoría</h3>
              <p className="text-xs text-slate-500 leading-relaxed mb-6">Relación cruzada de productos ordenados alfabéticamente por línea.</p>
            </div>
            <div className="grid grid-cols-2 gap-2 pt-3 border-t border-slate-100">
              <button
                onClick={rep3PDF}
                className="bg-gradient-to-r from-rose-500 to-rose-600 hover:from-rose-600 hover:to-rose-700 text-white font-bold text-xs py-2.5 rounded-xl shadow-sm transition"
              >
                📄 PDF
              </button>
              <button
                onClick={rep3Excel}
                className="bg-gradient-to-r from-emerald-600 to-teal-700 hover:from-emerald-700 hover:to-teal-800 text-white font-bold text-xs py-2.5 rounded-xl shadow-sm transition"
              >
                📊 Excel
              </button>
            </div>
          </div>

          {/* Tarjeta 4 */}
          <div className="bg-white rounded-2xl p-6 border border-slate-200/80 shadow-sm hover:shadow-md transition-all flex flex-col justify-between group">
            <div>
              <div className="flex items-center justify-between mb-3">
                <span className="bg-amber-50 text-amber-700 border border-amber-200 text-[10px] font-black tracking-wider uppercase px-2.5 py-1 rounded-md">
                  REPORTE 04
                </span>
                <span className="text-2xl group-hover:scale-110 transition-transform">⚠️</span>
              </div>
              <h3 className="font-bold text-slate-900 text-base mb-1">Medicamentos Stock Bajo</h3>
              <p className="text-xs text-slate-500 leading-relaxed mb-6">Alerta de bajo inventario con resaltado automático en Excel.</p>
            </div>
            <div className="pt-3 border-t border-slate-100">
              <button
                onClick={() => dialogStockRef.current?.showModal()}
                className="w-full bg-slate-900 hover:bg-slate-800 text-white font-bold text-xs py-2.5 rounded-xl shadow-sm transition flex items-center justify-center gap-2"
              >
                ⚙️ Configurar Umbral
              </button>
            </div>
          </div>

          {/* Tarjeta 5 */}
          <div className="bg-white rounded-2xl p-6 border border-slate-200/80 shadow-sm hover:shadow-md transition-all flex flex-col justify-between group">
            <div>
              <div className="flex items-center justify-between mb-3">
                <span className="bg-teal-50 text-teal-700 border border-teal-200 text-[10px] font-black tracking-wider uppercase px-2.5 py-1 rounded-md">
                  REPORTE 05
                </span>
                <span className="text-2xl group-hover:scale-110 transition-transform">👥</span>
              </div>
              <h3 className="font-bold text-slate-900 text-base mb-1">Ventas por Cliente</h3>
              <p className="text-xs text-slate-500 leading-relaxed mb-6">Resumen general de facturación activa y montos por comprador.</p>
            </div>
            <div className="grid grid-cols-2 gap-2 pt-3 border-t border-slate-100">
              <button
                onClick={rep5PDF}
                className="bg-gradient-to-r from-rose-500 to-rose-600 hover:from-rose-600 hover:to-rose-700 text-white font-bold text-xs py-2.5 rounded-xl shadow-sm transition"
              >
                📄 PDF
              </button>
              <button
                onClick={rep5Excel}
                className="bg-gradient-to-r from-emerald-600 to-teal-700 hover:from-emerald-700 hover:to-teal-800 text-white font-bold text-xs py-2.5 rounded-xl shadow-sm transition"
              >
                📊 Excel
              </button>
            </div>
          </div>

          {/* Tarjeta 6 */}
          <div className="bg-white rounded-2xl p-6 border border-slate-200/80 shadow-sm hover:shadow-md transition-all flex flex-col justify-between group">
            <div>
              <div className="flex items-center justify-between mb-3">
                <span className="bg-teal-50 text-teal-700 border border-teal-200 text-[10px] font-black tracking-wider uppercase px-2.5 py-1 rounded-md">
                  REPORTE 06
                </span>
                <span className="text-2xl group-hover:scale-110 transition-transform">📋</span>
              </div>
              <h3 className="font-bold text-slate-900 text-base mb-1">Ventas Cliente Detalle</h3>
              <p className="text-xs text-slate-500 leading-relaxed mb-6">Unión completa de 4 tablas relacionales con fórmula de total general.</p>
            </div>
            <div className="grid grid-cols-2 gap-2 pt-3 border-t border-slate-100">
              <button
                onClick={rep6PDF}
                className="bg-gradient-to-r from-rose-500 to-rose-600 hover:from-rose-600 hover:to-rose-700 text-white font-bold text-xs py-2.5 rounded-xl shadow-sm transition"
              >
                📄 PDF
              </button>
              <button
                onClick={rep6Excel}
                className="bg-gradient-to-r from-emerald-600 to-teal-700 hover:from-emerald-700 hover:to-teal-800 text-white font-bold text-xs py-2.5 rounded-xl shadow-sm transition"
              >
                📊 Excel
              </button>
            </div>
          </div>

        </div>
      </main>

      {/* DIÁLOGO MODAL REPORTE 2 */}
      <dialog
        ref={dialogNombreRef}
        className="rounded-2xl p-6 shadow-2xl backdrop:bg-slate-900/60 border border-slate-200 max-w-md w-full bg-white"
      >
        <div className="flex items-center justify-between mb-4 border-b pb-3 border-slate-100">
          <h3 className="font-bold text-slate-900 text-base">Filtrar por Nombre</h3>
          <span className="text-[10px] font-bold bg-teal-50 text-teal-700 px-2 py-0.5 rounded border border-teal-200">
            Reporte 02
          </span>
        </div>
        <p className="text-xs text-slate-500 mb-4">Ingrese un texto para buscar en el catálogo (Ejemplo: crema, acetaminofén):</p>
        <input
          type="text"
          value={filtroNombre}
          onChange={(e) => setFiltroNombre(e.target.value)}
          placeholder="Ej: crema"
          className="border border-slate-300 p-2.5 rounded-xl w-full mb-6 text-sm outline-none focus:ring-2 focus:ring-teal-600 focus:border-transparent"
        />
        <div className="flex gap-2 justify-end pt-3 border-t border-slate-100">
          <button
            onClick={() => dialogNombreRef.current?.close()}
            className="px-4 py-2 text-slate-500 text-xs font-semibold hover:bg-slate-100 rounded-xl transition"
          >
            Cancelar
          </button>
          <button
            onClick={rep2PDF}
            className="bg-rose-600 hover:bg-rose-700 text-white font-bold text-xs px-4 py-2 rounded-xl transition"
          >
            Exportar PDF
          </button>
          <button
            onClick={rep2Excel}
            className="bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs px-4 py-2 rounded-xl transition"
          >
            Exportar Excel
          </button>
        </div>
      </dialog>

      {/* DIÁLOGO MODAL REPORTE 4 */}
      <dialog
        ref={dialogStockRef}
        className="rounded-2xl p-6 shadow-2xl backdrop:bg-slate-900/60 border border-slate-200 max-w-md w-full bg-white"
      >
        <div className="flex items-center justify-between mb-4 border-b pb-3 border-slate-100">
          <h3 className="font-bold text-slate-900 text-base">Filtrar Stock Bajo</h3>
          <span className="text-[10px] font-bold bg-amber-50 text-amber-700 px-2 py-0.5 rounded border border-amber-200">
            Reporte 04
          </span>
        </div>
        <p className="text-xs text-slate-500 mb-4">Ingrese la cantidad máxima de existencias para filtrar:</p>
        <input
          type="number"
          value={filtroStock}
          onChange={(e) => setFiltroStock(Number(e.target.value))}
          className="border border-slate-300 p-2.5 rounded-xl w-full mb-6 text-sm outline-none focus:ring-2 focus:ring-teal-600 focus:border-transparent"
        />
        <div className="flex gap-2 justify-end pt-3 border-t border-slate-100">
          <button
            onClick={() => dialogStockRef.current?.close()}
            className="px-4 py-2 text-slate-500 text-xs font-semibold hover:bg-slate-100 rounded-xl transition"
          >
            Cancelar
          </button>
          <button
            onClick={rep4PDF}
            className="bg-rose-600 hover:bg-rose-700 text-white font-bold text-xs px-4 py-2 rounded-xl transition"
          >
            Exportar PDF
          </button>
          <button
            onClick={rep4Excel}
            className="bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs px-4 py-2 rounded-xl transition"
          >
            Exportar Excel
          </button>
        </div>
      </dialog>
    </div>
  );
};