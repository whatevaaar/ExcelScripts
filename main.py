import xlrd, unidecode
from producto import Producto as Producto
import utilidadesHTML

RUTA_EXCEL = "db.xlsx"
RUTA_PRODUCTOS_HTML = "C:\\Users\\emili\\Documents\\Code\\Web\\HA\\productos\\"
RUTA_CATALOGO = "C:\\Users\\emili\\Documents\\Code\\Web\\HA\\catalogo.html"


def crearProductosDeHojaExcel(hojaExcel):
    listaDeProductos = []
    descripcionPrevia = ""
    for rowEx in range(1, hojaExcel.nrows):
        titulo = str(hojaExcel.cell(rowEx, 2).value)
        isbn = str(hojaExcel.cell(rowEx, 0).value)
        if (not isbn or not titulo):  # Tenemos un espacio vacío o un campo en blanco
            continue
        marca = str(hojaExcel.cell(rowEx, 3).value)
        area = str(hojaExcel.cell(rowEx, 4).value)
        sector = str(hojaExcel.cell(rowEx, 1).value)
        descripcion = str(hojaExcel.cell(rowEx, 5).value)
        if not descripcion:
            descripcion = descripcionPrevia
        else:
            descripcionPrevia = descripcion
        precio = '$' + str(hojaExcel.cell(rowEx, 6).value)
        listaDeProductos.append(Producto(titulo, isbn, precio, marca, area, sector, descripcion))
    return listaDeProductos

def crearProductosDeHojaExcelReaders(hojaExcel):
    listaDeProductos = []
    marca = "e-future"
    sector = "Readers"
    for rowEx in range(1, hojaExcel.nrows):
        titulo = str(hojaExcel.cell(rowEx, 1).value)
        isbn = str(hojaExcel.cell(rowEx, 0).value)
        if (not isbn or not titulo):  # Tenemos un espacio vacío o un campo en blanco
            continue
        area = str(hojaExcel.cell(rowEx, 2).value)
        componente = str(hojaExcel.cell(rowEx, 3).value)
        descripcion = titulo + " incluye " + componente \
            if (componente != "-") else titulo
        precio = '$' + str(hojaExcel.cell(rowEx, 5).value)
        listaDeProductos.append(Producto(titulo, isbn, precio, marca, area, sector, descripcion))
    return listaDeProductos

def regresarDescripcionDeLiteratura(area):
    if "cantar" in area:
        return "Incluye cd para el maestro"
    elif "Young" in area:
        return "Todas las obras de la Serie Young Readers " \
               "cuentan con su respectivo cuaderno de actividades y cd de audio"
    elif "Puertitas" in area:
        return "Todas las obras de esta subserie cuentan con un " \
               "cd multimedia que contiene todas las obras"
    elif "Puertas" in area:
        return "Todas las obras de la Serie Puertas cuentan con " \
               "su respectivo cuaderno de actividades"
    else:
        return ""

def crearProductosDeHojaExcelLiteratura(hojaExcel):
    listaDeProductos = []
    sector = "Literatura"
    areaPrevia = ""
    for rowEx in range(1, hojaExcel.nrows):
        titulo = str(hojaExcel.cell(rowEx, 1).value)
        isbn = str(hojaExcel.cell(rowEx, 0).value)
        if (not isbn and not titulo):  # Tenemos un espacio vacío o un campo en blanco
            continue
        if (titulo and not isbn):
            areaPrevia = titulo
            continue
        marca = ""
        area = areaPrevia
        autor = str(hojaExcel.cell(rowEx, 2).value)
        descripcion = "La obra de " + titulo + "por" + autor + "." + regresarDescripcionDeLiteratura(area)
        precio = '$' + str(hojaExcel.cell(rowEx, 3).value)
        listaDeProductos.append(Producto(titulo, isbn, precio, marca, area, sector, descripcion))
    return listaDeProductos

def escribirDumpCarrusel(listaDeProductos):
    listaDeProductosEnCarrusel = []
    with open('dumpCarrusel.html', 'w') as htmlCarrusel:
        for producto in listaDeProductos:
            productoEnCarrusel =utilidadesHTML.crearElementoCarruselHTMLDeProducto(producto)
            listaDeProductosEnCarrusel.append(productoEnCarrusel)
            htmlCarrusel.write( productoEnCarrusel )
    return listaDeProductosEnCarrusel

def escribirDumpGrid(listaDeProductos):
    listaDeProductosEnGrid = []
    with open('dumpGrid.html', 'w') as htmlGrid:
        for producto in listaDeProductos:
            productoEnGrid =utilidadesHTML.crearElementoGridHTMLDeProducto(producto)
            listaDeProductosEnGrid.append(productoEnGrid)
            htmlGrid.write( productoEnGrid )
    return listaDeProductosEnGrid

def escribirHojasHTML(listaDeProductos, listaDeProductosEnCarrusel):
    for producto in listaDeProductos:
        with open(RUTA_PRODUCTOS_HTML + producto.enlace + ".html", 'wb') as productoHtml:
                productoHtml.write(utilidadesHTML.crearHojaHTMLDeProducto(producto, listaDeProductosEnCarrusel) )

def escribirCatalogo(listaDeProductosEnGrid):
    with open(RUTA_CATALOGO,'wb') as catalogo:
        catalogo.write( utilidadesHTML.crearCatalogoHTML( listaDeProductosEnGrid ) )

if __name__ == '__main__':
    listaDeProductosEnGrid = []
    listaDeProductosEnCarrusel = []
    productos = []
    workbook = xlrd.open_workbook(RUTA_EXCEL)
    preciosEsp = workbook.sheet_by_index(0)
    preciosLiteratura = workbook.sheet_by_index(1)
    preciosELT = workbook.sheet_by_index(2)
    preciosReaders = workbook.sheet_by_index(3)

    productos.extend( crearProductosDeHojaExcel(preciosELT) )
    productos.extend( crearProductosDeHojaExcel(preciosELT) )
    productos.extend( crearProductosDeHojaExcelLiteratura(preciosLiteratura) )
    productos.extend( crearProductosDeHojaExcelReaders(preciosReaders) )
    listaDeProductosEnCarrusel = escribirDumpCarrusel(productos) 
    listaDeProductosEnGrid = escribirDumpGrid(productos) 
    escribirHojasHTML(productos,listaDeProductosEnCarrusel)
    #escribirCatalogo( listaDeProductosEnGrid )

