export interface CategoriaDTO {
  idCategoria: number;
  estado: boolean;
  nombre: string;
  descripcion: string;
}

export interface Medicamento {
  idMedicamento: number;
  estado: boolean;
  nombre: string;
  laboratorio: string;
  precio: number;
  stock: number;
}

export interface MedicamentoCategoriaDTO {
  nombreMedicamento: string;
  laboratorio: string;
  nombreCategoria: string;
  precio: number;
  stock: number;
}

export interface ClienteVentaDTO {
  nombreCliente: string;
  apellidoCliente: string;
  fechaVenta: string;
  total: number;
}

export interface ClienteVentaDetalleDTO {
  nombreCliente: string;
  apellidoCliente: string;
  fechaVenta: string;
  nombreMedicamento: string;
  cantidad: number;
  precioUnitario: number;
  subtotal: number;
}