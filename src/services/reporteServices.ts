import api from '../api/axios';
import type { 
  CategoriaDTO, 
  Medicamento, 
  MedicamentoCategoriaDTO, 
  ClienteVentaDTO, 
  ClienteVentaDetalleDTO 
} from '../types/reportes';

// 1. Categorías Activas
export const getCategoriasActivas = () => 
  api.get<CategoriaDTO[]>('/categorias/mostrarActivos');

// 2. Medicamentos por Filtro Nombre
export const getMedicamentosFiltro = (nombre: string) => 
  api.get<Medicamento[]>(`/medicamentos/mostrarActivosFiltro?nombre=${nombre}`);

// 3. Medicamentos con su Categoría
export const getReporteMedicamentoCategoria = () => 
  api.get<MedicamentoCategoriaDTO[]>('/medicamentos/reporteMedicamentoCategoria');

// 4. Medicamentos Stock Bajo
export const getMedicamentosStockBajo = (cantidad: number) => 
  api.get<Medicamento[]>(`/medicamentos/stockBajo?cantidad=${cantidad}`);

// 5. Ventas por Cliente
export const getReporteClienteVenta = () => 
  api.get<ClienteVentaDTO[]>('/clientes/reporteClienteVenta');

// 6. Detalle de Ventas por Cliente
export const getReporteClienteVentaDetalle = () => 
  api.get<ClienteVentaDetalleDTO[]>('/clientes/reporteClienteVentaDetalle');