import unidecode

class Producto:
   def __init__(self,titulo,isbn,precio,marca,area, sector, descripcion):
       self.titulo = titulo
       self.isbn = isbn
       self.precio = precio
       self.marca = marca
       self.area = area
       self.area = area
       self.sector = sector
       self.descripcion = descripcion
       self.descripcionCorta = self.crearDescripcionCorta()
       self.enlace = self.crearEnlace()
       self.enlaceHTMLRelativo = self.crearEnlaceHTMLRelativo()
       self.enlaceIMGRelativo = self.crearEnlaceIMGRelativo()
       self.enlaceHTMLAbsoluto = self.crearEnlaceHTMLAbsoluto()
       self.enlaceIMGAbsoluto = self.crearEnlaceIMGAbsoluto()

   def crearEnlaceHTMLAbsoluto(self):
       return "productos/" + self.enlace + ".html"

   def crearEnlaceIMGAbsoluto(self):
       return "img/productos/" + self.enlace + ".jpg"
   def crearEnlaceHTMLRelativo(self):
       return "productos/" + self.enlace + ".html"

   def crearEnlaceIMGRelativo(self):
       return "../img/productos/" + self.enlace + ".jpg"

   def crearDescripcionCorta(self):
        MAX_PALABRAS = 20
        descripcionCorta = []
        for contador, palabra in enumerate(self.descripcion.split()):
            if contador <= MAX_PALABRAS:
                descripcionCorta.append(palabra)
        return ' '.join(descripcionCorta) + "..."

   def crearEnlace(self):
        enlace = self.titulo
        enlace = unidecode.unidecode(enlace)\
            .replace(' ', '-').replace('?','')\
            .replace(':', '').replace('.', '')\
            .replace('/', '-').replace('*','')
        return enlace
