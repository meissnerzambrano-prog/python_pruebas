"""Inventario de productos.

Registra productos con cantidad, precio de compra y precio de venta, calcula
ganancias y exporta los datos a TXT o CSV (se abre en Excel).
Los datos se guardan solos en 'productos.json', junto a este archivo.
"""
import csv
import json
import math
import os
import re
import sys
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

CARPETA = os.path.dirname(sys.executable if getattr(sys, "frozen", False)
                          else os.path.abspath(__file__))
ARCHIVO_DATOS = os.path.join(CARPETA, "productos.json")


# ---------------------------------------------------------------- números
def leer_numero(texto):
    """Convierte '12.500', '12500', '12,5' o '$ 1.250,50' en número.
    Lanza ValueError si no es válido o es negativo."""
    t = texto.strip().replace("$", "").replace(" ", "")
    if not t:
        raise ValueError("vacío")
    if "." in t and "," in t:                       # 1.250,50 -> punto = miles
        t = t.replace(".", "").replace(",", ".")
    elif "," in t:
        if re.fullmatch(r"\d{1,3}(,\d{3})+", t):    # 1,250 -> miles
            t = t.replace(",", "")
        else:                                       # 12,5 -> decimal
            t = t.replace(",", ".")
    elif "." in t and re.fullmatch(r"\d{1,3}(\.\d{3})+", t):
        t = t.replace(".", "")                      # 12.500 -> miles
    valor = float(t)
    if not math.isfinite(valor) or valor < 0:
        raise ValueError("número no válido")
    return valor


def entero(n):
    return f"{int(n):,}".replace(",", ".")


def dinero(n):
    """12500 -> $12.500   |   1250.5 -> $1.250,50"""
    signo = "-" if n < 0 else ""
    n = abs(n)
    s = f"{round(n):,}" if abs(n - round(n)) < 0.005 else f"{n:,.2f}"
    return signo + "$" + s.replace(",", "X").replace(".", ",").replace("X", ".")


def porcentaje(p):
    if not p["compra"]:
        return "-"
    gan = (p["venta"] - p["compra"]) / p["compra"] * 100
    return f"{gan:.1f}%".replace(".", ",")


def numero_texto(n):
    return str(int(n)) if n == int(n) else str(n)


# ---------------------------------------------------------------- cálculos
def calcular(p):
    """% de ganancia = ganancia por unidad / precio de compra."""
    gan_u = p["venta"] - p["compra"]
    return {
        "ganancia_unidad": gan_u,
        "margen": gan_u / p["compra"] * 100 if p["compra"] else 0.0,
        "inversion": p["compra"] * p["cantidad"],
        "ventas": p["venta"] * p["cantidad"],
        "ganancia_total": gan_u * p["cantidad"],
    }


def totales(productos):
    t = {"productos": len(productos), "unidades": 0,
         "inversion": 0.0, "ventas": 0.0, "ganancia": 0.0}
    for p in productos:
        c = calcular(p)
        t["unidades"] += p["cantidad"]
        t["inversion"] += c["inversion"]
        t["ventas"] += c["ventas"]
        t["ganancia"] += c["ganancia_total"]
    return t


# ---------------------------------------------------------------- archivos
def guardar_datos(productos, ruta=ARCHIVO_DATOS):
    temporal = ruta + ".tmp"            # se escribe aparte para no dañar el original
    with open(temporal, "w", encoding="utf-8") as f:
        json.dump(productos, f, ensure_ascii=False, indent=1)
    os.replace(temporal, ruta)


def cargar_datos(ruta=ARCHIVO_DATOS):
    if not os.path.exists(ruta):
        return []
    with open(ruta, encoding="utf-8") as f:
        return json.load(f)


def exportar_txt(productos, ruta):
    ancho = 30
    enc = (f"{'Cód':>5}  {'Producto':<{ancho}}  {'Cant.':>7}  {'P. compra':>13}  "
           f"{'P. venta':>13}  {'Gan./unid.':>13}  {'% gan.':>8}  {'Gan. total':>15}")
    sep = "-" * len(enc)
    t = totales(productos)
    lineas = ["REPORTE DE PRODUCTOS", f"Generado: {datetime.now():%Y-%m-%d %H:%M}",
              "", sep, enc, sep]
    for p in productos:
        c = calcular(p)
        lineas.append(
            f"{p['id']:>5}  {p['nombre'][:ancho]:<{ancho}}  {entero(p['cantidad']):>7}  "
            f"{dinero(p['compra']):>13}  {dinero(p['venta']):>13}  "
            f"{dinero(c['ganancia_unidad']):>13}  {porcentaje(p):>8}  "
            f"{dinero(c['ganancia_total']):>15}")
    lineas += [sep, "", "RESUMEN",
               f"Productos registrados : {entero(t['productos'])}",
               f"Unidades en total     : {entero(t['unidades'])}",
               f"Inversión (compra)    : {dinero(t['inversion'])}",
               f"Ventas esperadas      : {dinero(t['ventas'])}",
               f"Ganancia total        : {dinero(t['ganancia'])}"]
    with open(ruta, "w", encoding="utf-8") as f:
        f.write("\n".join(lineas) + "\n")


def exportar_csv(productos, ruta):
    """CSV con ';' para que Excel en español lo separe en columnas."""
    def num(n):
        return str(int(n)) if n == int(n) else str(round(n, 2)).replace(".", ",")

    with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Código", "Producto", "Cantidad", "Precio de compra", "Precio de venta",
                    "Ganancia por unidad", "% ganancia", "Inversión", "Ventas esperadas",
                    "Ganancia total"])
        for p in productos:
            c = calcular(p)
            w.writerow([p["id"], p["nombre"], p["cantidad"], num(p["compra"]), num(p["venta"]),
                        num(c["ganancia_unidad"]), num(round(c["margen"], 1)),
                        num(c["inversion"]), num(c["ventas"]), num(c["ganancia_total"])])


# ---------------------------------------------------------------- ventana
COLUMNAS = [  # clave, título, ancho, alineación
    ("id", "Cód.", 60, "e"), ("nombre", "Producto", 260, "w"), ("cantidad", "Cantidad", 80, "e"),
    ("compra", "P. compra", 110, "e"), ("venta", "P. venta", 110, "e"),
    ("ganancia_unidad", "Gan./unid.", 110, "e"), ("margen", "% gan.", 70, "e"),
    ("ganancia_total", "Gan. total", 130, "e"),
]


def valor_orden(p, clave):
    if clave == "nombre":
        return p["nombre"].lower()
    if clave in ("id", "cantidad", "compra", "venta"):
        return p[clave]
    return calcular(p)[clave]


class Inventario:
    def __init__(self, ventana):
        self.ventana = ventana
        ventana.title("Inventario de productos")
        ventana.geometry("1020x660")
        ventana.minsize(900, 540)

        self.productos = self.cargar()
        self.siguiente_id = max((p["id"] for p in self.productos), default=0) + 1
        self.editando = None
        self.orden = ("id", False)

        self.v_nombre, self.v_cantidad = tk.StringVar(), tk.StringVar()
        self.v_compra, self.v_venta = tk.StringVar(), tk.StringVar()
        self.v_buscar, self.v_estado = tk.StringVar(), tk.StringVar()
        self.v_totales = {k: tk.StringVar() for k in
                          ("productos", "unidades", "inversion", "ventas", "ganancia")}
        self.construir()
        self.refrescar()

    # --- carga inicial
    def cargar(self):
        try:
            return cargar_datos()
        except (OSError, ValueError):
            copia = ARCHIVO_DATOS.replace(".json", f"_danado_{datetime.now():%Y%m%d_%H%M%S}.json")
            try:
                os.replace(ARCHIVO_DATOS, copia)
            except OSError:
                pass
            messagebox.showwarning("Archivo dañado",
                                   "No pude leer productos.json. Guardé una copia aparte "
                                   "y empiezo con la lista vacía.")
            return []

    # --- construcción de la ventana
    def construir(self):
        v = self.ventana
        v.columnconfigure(0, weight=1)
        v.rowconfigure(2, weight=1)

        form = ttk.LabelFrame(v, text=" Producto ", padding=10)
        form.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        campos = [("Nombre", self.v_nombre, 34), ("Cantidad", self.v_cantidad, 10),
                  ("Precio de compra", self.v_compra, 14), ("Precio de venta", self.v_venta, 14)]
        self.entradas = []
        for i, (texto, var, ancho) in enumerate(campos):
            ttk.Label(form, text=texto).grid(row=0, column=i, sticky="w")
            e = ttk.Entry(form, textvariable=var, width=ancho)
            e.grid(row=1, column=i, sticky="w", padx=(0, 10))
            e.bind("<Return>", lambda _e: self.guardar_formulario())
            self.entradas.append(e)
        botones = ttk.Frame(form)
        botones.grid(row=1, column=4, padx=(6, 0))
        ttk.Button(botones, text="Agregar", command=self.agregar).pack(side="left", padx=2)
        self.btn_actualizar = ttk.Button(botones, text="Actualizar", state="disabled",
                                         command=self.actualizar)
        self.btn_actualizar.pack(side="left", padx=2)
        self.btn_eliminar = ttk.Button(botones, text="Eliminar", state="disabled",
                                       command=self.eliminar)
        self.btn_eliminar.pack(side="left", padx=2)
        ttk.Button(botones, text="Limpiar", command=self.limpiar).pack(side="left", padx=2)

        barra = ttk.Frame(v)
        barra.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        ttk.Label(barra, text="Buscar:").pack(side="left")
        ttk.Entry(barra, textvariable=self.v_buscar, width=32).pack(side="left", padx=6)
        self.v_buscar.trace_add("write", lambda *_: self.refrescar())
        ttk.Button(barra, text="Exportar CSV (Excel)",
                   command=lambda: self.exportar("csv")).pack(side="right", padx=(6, 0))
        ttk.Button(barra, text="Exportar TXT",
                   command=lambda: self.exportar("txt")).pack(side="right")

        marco = ttk.Frame(v)
        marco.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        marco.columnconfigure(0, weight=1)
        marco.rowconfigure(0, weight=1)
        ttk.Style().configure("Treeview", rowheight=24)
        self.tabla = ttk.Treeview(marco, columns=[c[0] for c in COLUMNAS],
                                  show="headings", selectmode="browse")
        for clave, texto, ancho, ancla in COLUMNAS:
            self.tabla.heading(clave, text=texto, command=lambda k=clave: self.ordenar(k))
            self.tabla.column(clave, width=ancho, anchor=ancla, stretch=(clave == "nombre"))
        scroll = ttk.Scrollbar(marco, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scroll.set)
        self.tabla.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        self.tabla.tag_configure("perdida", foreground="#c0392b")
        self.tabla.bind("<<TreeviewSelect>>", self.seleccionar)
        self.tabla.bind("<Delete>", lambda _e: self.eliminar())

        resumen = ttk.Frame(v)
        resumen.grid(row=3, column=0, sticky="ew", padx=10, pady=5)
        for i, (clave, titulo) in enumerate([("productos", "Productos"), ("unidades", "Unidades"),
                                             ("inversion", "Inversión"), ("ventas", "Ventas esperadas"),
                                             ("ganancia", "Ganancia total")]):
            resumen.columnconfigure(i, weight=1)
            ttk.Label(resumen, text=titulo, foreground="gray").grid(row=0, column=i, sticky="w")
            ttk.Label(resumen, textvariable=self.v_totales[clave],
                      font=("Arial", 13, "bold")).grid(row=1, column=i, sticky="w")
        ttk.Label(v, textvariable=self.v_estado, foreground="gray").grid(
            row=4, column=0, sticky="w", padx=10, pady=(0, 8))

    # --- mostrar datos
    def refrescar(self):
        q = self.v_buscar.get().strip().lower()
        filas = [p for p in self.productos
                 if not q or q in p["nombre"].lower() or q == str(p["id"])]
        clave, desc = self.orden
        filas.sort(key=lambda p: valor_orden(p, clave), reverse=desc)

        self.tabla.delete(*self.tabla.get_children())
        for p in filas:
            c = calcular(p)
            self.tabla.insert("", "end", iid=str(p["id"]),
                              tags=("perdida",) if c["ganancia_unidad"] < 0 else (),
                              values=(p["id"], p["nombre"], entero(p["cantidad"]),
                                      dinero(p["compra"]), dinero(p["venta"]),
                                      dinero(c["ganancia_unidad"]), porcentaje(p),
                                      dinero(c["ganancia_total"])))
        for k, texto, *_ in COLUMNAS:
            flecha = (" ▼" if desc else " ▲") if k == clave else ""
            self.tabla.heading(k, text=texto + flecha)

        t = totales(self.productos)
        self.v_totales["productos"].set(entero(t["productos"]))
        self.v_totales["unidades"].set(entero(t["unidades"]))
        self.v_totales["inversion"].set(dinero(t["inversion"]))
        self.v_totales["ventas"].set(dinero(t["ventas"]))
        self.v_totales["ganancia"].set(dinero(t["ganancia"]))
        self.v_estado.set(f"Mostrando {len(filas)} de {len(self.productos)} productos. "
                          "Las filas en rojo se venden por menos de lo que costaron.")

    def ordenar(self, clave):
        self.orden = (clave, not self.orden[1]) if self.orden[0] == clave else (clave, False)
        self.refrescar()

    # --- formulario
    def buscar(self, id_):
        return next(p for p in self.productos if p["id"] == id_)

    def seleccionar(self, _evento=None):
        sel = self.tabla.selection()
        if not sel:
            return
        p = self.buscar(int(sel[0]))
        self.editando = p["id"]
        self.v_nombre.set(p["nombre"])
        self.v_cantidad.set(str(p["cantidad"]))
        self.v_compra.set(numero_texto(p["compra"]))
        self.v_venta.set(numero_texto(p["venta"]))
        self.btn_actualizar.config(state="normal")
        self.btn_eliminar.config(state="normal")

    def limpiar(self):
        self.editando = None
        for var in (self.v_nombre, self.v_cantidad, self.v_compra, self.v_venta):
            var.set("")
        self.tabla.selection_remove(self.tabla.selection())
        self.btn_actualizar.config(state="disabled")
        self.btn_eliminar.config(state="disabled")
        self.entradas[0].focus_set()

    def aviso(self, mensaje, campo):
        messagebox.showwarning("Revisa los datos", mensaje)
        self.entradas[campo].focus_set()

    def leer_formulario(self):
        nombre = self.v_nombre.get().strip()
        if not nombre:
            return self.aviso("Escribe el nombre del producto.", 0)
        try:
            cantidad = leer_numero(self.v_cantidad.get())
        except ValueError:
            return self.aviso("La cantidad debe ser un número, por ejemplo 25.", 1)
        if cantidad != int(cantidad):
            return self.aviso("La cantidad debe ser un número entero.", 1)
        try:
            compra = leer_numero(self.v_compra.get())
        except ValueError:
            return self.aviso("El precio de compra debe ser un número, por ejemplo 12500.", 2)
        try:
            venta = leer_numero(self.v_venta.get())
        except ValueError:
            return self.aviso("El precio de venta debe ser un número, por ejemplo 15000.", 3)
        if venta < compra and not messagebox.askyesno(
                "Precio de venta bajo",
                "El precio de venta es menor al de compra: perderías dinero.\n"
                "¿Guardar de todas formas?"):
            return None
        return nombre[:80], int(cantidad), compra, venta

    def guardar_formulario(self):
        self.actualizar() if self.editando is not None else self.agregar()

    def guardar(self):
        try:
            guardar_datos(self.productos)
        except OSError as e:
            messagebox.showerror("No se pudo guardar", f"No pude escribir productos.json:\n{e}")

    # --- acciones
    def agregar(self):
        datos = self.leer_formulario()
        if not datos:
            return
        nombre, cantidad, compra, venta = datos
        if any(p["nombre"].lower() == nombre.lower() for p in self.productos) and \
                not messagebox.askyesno("Producto repetido",
                                        "Ya existe un producto con ese nombre.\n"
                                        "¿Agregarlo de todas formas?"):
            return
        nuevo = {"id": self.siguiente_id, "nombre": nombre, "cantidad": cantidad,
                 "compra": compra, "venta": venta}
        self.siguiente_id += 1
        self.productos.append(nuevo)
        self.guardar()
        self.limpiar()
        self.refrescar()
        if self.tabla.exists(str(nuevo["id"])):
            self.tabla.see(str(nuevo["id"]))
        self.v_estado.set(f"Agregado: {nombre}")

    def actualizar(self):
        if self.editando is None:
            return
        datos = self.leer_formulario()
        if not datos:
            return
        p = self.buscar(self.editando)
        p["nombre"], p["cantidad"], p["compra"], p["venta"] = datos
        self.guardar()
        self.limpiar()
        self.refrescar()
        self.v_estado.set(f"Actualizado: {p['nombre']}")

    def eliminar(self):
        if self.editando is None:
            return
        p = self.buscar(self.editando)
        if messagebox.askyesno("Eliminar producto", f"¿Eliminar «{p['nombre']}»?"):
            self.productos.remove(p)
            self.guardar()
            self.limpiar()
            self.refrescar()
            self.v_estado.set(f"Eliminado: {p['nombre']}")

    def exportar(self, tipo):
        if not self.productos:
            messagebox.showinfo("Sin datos", "Todavía no hay productos para exportar.")
            return
        ext = ".txt" if tipo == "txt" else ".csv"
        tipos = [("Archivo de texto", "*.txt")] if tipo == "txt" else [("CSV (Excel)", "*.csv")]
        ruta = filedialog.asksaveasfilename(
            title="Guardar archivo", defaultextension=ext, filetypes=tipos,
            initialfile=f"productos_{datetime.now():%Y-%m-%d}{ext}")
        if not ruta:
            return
        try:
            (exportar_txt if tipo == "txt" else exportar_csv)(self.productos, ruta)
        except OSError as e:
            messagebox.showerror("No se pudo guardar",
                                 f"{e}\n\nSi el archivo está abierto en Excel, ciérralo e intenta de nuevo.")
        else:
            messagebox.showinfo("Listo", f"Archivo guardado en:\n{ruta}")


if __name__ == "__main__":
    ventana = tk.Tk()
    Inventario(ventana)
    ventana.mainloop()