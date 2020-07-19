# AoYind 3 - Cliente

Importante, no bajar el codigo con el boton Download as a ZIP de github por que lo descarga mal, muchos archivos por el encoding quedan corruptos.

Tenes que bajar el codigo con un cliente de git, con el cliente original de la linea de comandos seria:
```
git clone https://github.com/YindSoft/aoyind3-client.git
```

## Como utilizar el cliente del juego.

En este repositorio solo se encuentra los codigos de fuente del cliente, por lo tanto para poder ejecutarlo correctamente es necesario tambien clonar el repo de resources para copiar los archivos necesarios.
Esto esta hecho así para poder separar bien los cambios de los recursos del juego en general y lo que es codigo.

Pueden clonar el repo de recursos desde aquí.
```
git clone https://github.com/YindSoft/aoyind3-resources.git
```


Para el cliente es necesario copiar los siguientes archivos/carpetas:
```
Recursos/*
INIT/*
MP3/*
Midi/*


Los mapas tienen que copiarlos de la carpeta maps a la carpeta recursos.
Maps/Mapa1.AO -> Recursos/Mapa1.Ao
Maps/Mapa2.AO -> Recursos/Mapa2.Ao
```


Para poder configurar una IP diferente a localhost en el juego busquen en el modulo Mod_General estas lineas:

```
If False Then 'ipx
    IpServidor = "ip publica"
Else
    IpServidor = frmMain.Client.LocalIP 'localhost
End If
```

Posiblemente se cambie en el futuro eso ya que esta hardcodeado para testing basicamente.


## F.A.Q:

#### Error - Al abrir el proyecto en Visual Basic 6 no puede cargar todas las dependencias:
Este es un error comun que les suele pasar a varias personas, esto es debido que el EOL del archivo esta corrupto.
Visual Basic 6 lee el .vbp en CLRF, hay varias formas de solucionarlo:

Opcion a:
Con Notepad++ cambiar el EOL del archivo a CLRF

Opcion b:
Abrir un editor de texto y reemplazar todos los `'\n'` por `'\r\n'`

--------------------------

