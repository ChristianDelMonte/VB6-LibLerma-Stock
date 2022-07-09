# VB6-LibLerma-Stock
Sistema para Actualizar precios de Productos de Tienda Nube para Libreria Lerma

Para entender el funcionamiento de este sistema hay que tener en cuenta lo siguiente:
La Libreria utiliza un sistema propio donde se manejan los precios, el stock, las ventas
y demas movimientos de la mercaderia propias del negocio.
Se utiliza este sistema para exportar todos los productos y precios correspondientes en
un archivo CSV que utilizaremos más adelante y el cual llamaremos "ORIGEN".

Paralelamente a esto la Libreria posee un E-commerce realizado en Tienda Nube que también
posee una lista de productos a la venta y sus respectivos precios y stock, el cual usamos
también para exportar su contenido en otro archivo CSV al cual llamaremos "DESTINO"

Ahora ingresamos en nuestro sistema VB6-LibLerma-Stock y utilizando una base de datos
propia (interna) importamos el CSV obtenida del ORIGEN.
Una vez importados los datos vamos a la sección "actualizar CSV de tienda nube" y
seleccionamos el archivo obtenido de DESTINO y creamos un nuevo archivo donde se guardarán
los datos de DESTINO con las actualizaciones pertinentes obtenidas de ORIGEN.
Ese nuevo archivo, en formato CSV, luego del proceso se importará nuevamente a Tienda
Nube para actualizar los datos que sufrieron cambios (precios, stock, etc)

Este proceso se realiza diariamente, toma solo 5 minutos, y permite actualizar todos
los productos de Tienda Nube obteniendo datos del sistema de origen del negocio evitando
asi el tener que actualizar manualmente en tienda nube producto por producto.

Además se incorporó a este sistema un método para llevar el control de cierta mercadería
que debia ser dada de baja manualmente. el mismo utiliza la base de datos interna obtenida
del CSV de ORIGEN. Genera la baja correspondiente con codigos y cantidades y permite en
un simple reporte imprimir los resutados para dejar constancia en fisico del movimiento.

PS: El único inconveniente encontrado en el proceso y que debe ser solucionado a mano
es el echo de que Visual Basic genera el archivo CSV por comas agregándole " comillas
a cada registro y que luego deben ser renovidas antes de poder importar dicho archivo
a Tienda Nube para actualizar. Como aclaración cabe decir que esas comillas se reemplazan
facilmente realiando un "reemplazar" en un simple block de notas y reemplazando dichas
comillas por nada mas que un espacio vacío y volviendo a guardar el archivo con el
block de notas.
