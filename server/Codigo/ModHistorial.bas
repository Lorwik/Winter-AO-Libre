Attribute VB_Name = "ModHistorial"
'***************************************************************************************
'MODULO HISTORIAL - TODOS LOS CAMBIOS EN PROGRAMACION SE ANOTAN AQUI.
'***************************************************************************************

'- Agregado sistema de cuentas [Funcional] <Lorwik>
'- Agregado sistema de quest [Funcional] <Lorwik>
'- Agregado Nombre del Mapa. [Funcional] <Lorwik>
'- Sustituido musica midi por MP3. [Funcional] <Lorwik>
'- Barra de progreso en cargando. [Funcional] <Lorwik>
'- Modificado la informacion de carga en el form cargando. [Funcional] <Lorwik>
'- Barra de progreso de experiencia. [Funcional] <Lorwik>
'- Agregado minimapa. [Funciona] <Lorwik>
'- Mejoracion del minimapa. [Funcional] <Lorwik>
'- Nombre de los objetos en el label del inventario. [Funcional] <Lorwik>
'- Barra en fuerza y agilidad. [Funcional] <Lorwik>
'- Eliminado el Sinfo.dat y la antigua carga de ip de servidores. [Funcional] <Lorwik>
'- Atributos limitados a 35 <Lorwik>
'- Agregado sistema de Clima. [Funcional] <Lorwik>
'- Eliminado el archivo Colores.dat [Funcional] <Lorwik>
'- Desbugeado la carga de Graficos.ind <Lorwik>
'- Agregado nombre de los npc abajo de sus cuerpos [Funcional] <Lorwik>
'- Adaptado el sistema de inventario de 0.13 [Funcional] <Lorwik>
'- Agregado Comercio Grafico [Funcional <Lorwik>
'- Agregado Boveda Grafica [Funcional] <Lorwik>
'- Consola transparente con deteccion de SO [Funcional] <Lorwik>
'- Se elimino la opcion de elegir hogar en crear personaje. Ahora siempre nacera en Ramx. [Funcional] <Lorwik>
'- Se elimino el asignar skills en la creacion de presonajes. Ahora se asigna dentro del juego. [Funcional] <Lorwik>
'- Agregado sistema de pasos segun la zona donde se encuentre. Pasto, Nieve, Desierto, Dungeon, etc... [Funcional] <Lorwik>
'- Apartir de los 10k de oro no se cae [Funcional] <Lorwik>
'- Ahora los skills se asignan desde estadisticas [Funcional] <Lorwik>
'- Los Newbies ganan 500 monedas de oro al pasar de nivel [Funcional] <Lorwik>
'- Agregado Cabezas Seleccionables en la creación de personajes [Funcional] <Lorwik>
'- El sistema de Screenshots fue reemplazado por uno mejor. [Funcional] <Lorwik>
'- Agregado Mensaje de bienvenida al entrar juego [Funcional] <Lorwik>
'- Lo que se habla se queda en la consola, y tambien los npc [Por testear] <Lorwik>
'- Cuando haces doble click en un item equipable del inventario, te lo equipas [Funcional] <Lorwik>
'- Agregado engine DX8 [Funcional] <Lorwik>
'- Los newbie no pierden los items de newbie al morir [Funcional] <Lorwik>
'- Agregado sistema realista de Clima [Funcional] <Lorwik>
'- Se sacaron los molestos mensajes de "No ves nada interesante" y "Has recuperado X de mana" <Lorwik>
'- Al morir se oscurece el mapa <Lorwik>
'- Al pegarle un npc nos dice el daño arriba de la cabeza y si falla tambien. [Funcional] <Lorwik>
'- Ahora el Motd lo muestra en la consola [Funcional] <Lorwik>
'- La fuerza y la agilidad va a subir siempre hasta 35 como maximo sin importar la base [Funcional] <Lorwik>
'- Se elimino el dado, ahora los skills se asignan manualmente [Funcional] <Lorwik>
'- Se agrego indicadores de modificadores segun la raza en la creacion de personaje [Funcional] <Lorwik>
'- Se facilito la subida de skills naturales [Funcional] <Lorwik>
'- Se elimino el sistema de lluvia <Lorwik>
'- Guardado de nombre de cuenta en el conectar [Funcional] <Lorwik>
'- El Carpintero y el Herrero puede elegir la cantidad a construir [Funcional] <Lorwik>
'- Haciendo click en las ventanas las podemos mover <Lorwik>
'- Cuando los npc tiran mas de 1 k de oro va a directo a la billetera, de lo contrario cae al suelo [Funcional] <Lorwik>
'- Agregado mapa del juego en la tecla Q <Lorwik>
'- Agregado macros de comandos configurables [Funcional] <Lorwik>
'- Cuando caminas sobre nieve o arena del desierto dejas huellas [Funcional] <Lorwik>
'- Se añadio Sincronizacion Vertical [Funcional] <Lorwik>
'- Ahora se puede configurar el nivel de precarga [Funcional] <Lorwik>
'- Se elimino el SendCMSTXT, ahora para hablar por chat de clan sera igual, solo que al pulsar supr se abrira el SendTXT con el comando para hablar por clan ya puesto [Funcional] <Lorwik>
'- Agregado sistema de global [Funcional] <Lorwik>
'- Agregado sistema de montura con velocidad, defensa y golpe [Funcional] <Lorwik>
'- Se ha optimizado el sistema de mañana, dia, tarde y noche <Lorwik>
'- Agregado sistema de puntos [Funcional] <Lorwik>
'- Agregado sistema de canjes por puntos [Funcional] <Lorwik>
'- Agregado clase Orco [Funcional] <Lorwik>
'- Agregado contador de tiempo restante de paralisis y invisibilidad [Funcional] <Lorwik>
'- Agregado AutoTorneos [Funcional] <Lorwik>
'- Eliminado sistema de Pretorianos <Lorwik>
'- Ahora los npc tendran X probabilidad de dropear objetos [Funcional] <Lorwik>
'- Se mejoro el comando /GM, agregando motivo de la consulta, sugerencias y reportes de bugs [Funcional] <Lorwik>
'- Se mejoro la carga de mapa del server, acelerando el tiempo de carga [Funcional] <Lorwik>
'- Se mejoro el sistema de canjes, ahora carga los items al iniciar el server y hay posibilidad de un 50% de descuento completando X Quest <Lorwik>
'- Se ha creado una base de seguridad. Macros, Speed Hack y Cheats noobs no funcionan [Funcional] <Lorwik>
'- Ahora es posible banear y desbanear las cuentas [Funcional] <Lorwik>
'- Agregado sistema de ranking con el comando "/RANK" [Funcional] <Lorwik>
'- Movimiento "realista". Al atacar fisicamente mueve escudo y arma, al tirar hechizo solo el arma y al trabajar mueve la herramienta de trabajo [Funcional] <Lorwik>
'- Agregado hechizos por area [Funcional] <Lorwik>
'- Agregado sistema de Macro Asistido desde el servidor [Funcional] <MaxTus>
'- Incorporé otro Enum para el manejo de más paquetes [Funcional] <Maxtus>
'- Se unio todas las clases trabajadoras en 1 sola llamado "Trabajador". [Funcional] <Lorwik>
'- Ya no se pueden tirar mas los items de newbie [Funcional+ <Lorwik>
'- Incorporé el sistema de evento de experiencia x2, el mismo se sortea cada hora y hay un 10% de que se active [Funcional] <Maxtus>
'- Arreglé el sistema de autoresu y autocurar e incorporé una función que detecta la cercanía del usuario a X Npc [Funcional] <Maxtus>
'- Ahora los Npcs Lanzan hechizos de la misma forma que los usuarios (Dicen Palabras magicas, mismos FXs) [Funcional] <Maxtus>
'- Añadido un intervalo para resucitar [Funcional] <Maxtus>
'- Completado el sistema de Titanes [Funcional} <Lorwik>
'- Aquellos items que no podamos usar apareceran en el inventario rojo <Lorwik>
'- Ahora las denuncias pueden ser desactivadas por los GM <Lorwik>
'- Ahora cuando tenemos las teclas configuradas y hablamos no se mueve el pj <Lorwik>
