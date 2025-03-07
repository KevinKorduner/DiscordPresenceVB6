# Explicación y Uso del Código

## Descripción General
Este programa permite la integración de **Discord Rich Presence** en una aplicación desarrollada en Visual Basic 6 (VB6). Se encarga de enviar información sobre el estado del usuario al perfil de Discord, mostrando detalles en su estado de actividad.

Código creado y liberado por **Kevin Korduner** el **07/03/2025** , actualmente es utilizado en **AOForever**.

---

## Características Principales
- Inicializa la conexión con **Discord Rich Presence** a través de la librería `discord-rpc.dll`.
- Actualiza el estado del usuario con un mensaje personalizado mientras juega **AOForever**. (Cambien el texto por su juego)
- Gestiona la memoria para enviar correctamente cadenas ANSI a la API de Discord.
- Ejecuta un temporizador que mantiene activa la comunicación con Discord.
- Cierra correctamente la conexión cuando el programa se cierra.

---

## Requisitos
1. **Discord** debe estar instalado y en ejecución en el sistema.
2. **Visual Basic 6** para compilar y ejecutar el programa.
3. La librería **discord-rpc.dll** debe estar en la misma carpeta que el ejecutable.

---

## Configuración y Uso
### 1. Configurar la App ID de Discord
La aplicación requiere una **App ID de Discord**, que se obtiene creando una aplicación en el https://discord.com/developers/applications

En el código, la App ID está definida en la función `Discord_Initialize`:

```vb
Discord_Initialize "5234441513530066311", handlers, 1, ""
```

Si deseas usar una App ID diferente, reemplaza **"5234441513530066311"** por la ID correspondiente a tu aplicación de Discord.

### 2. Ubicar `discord-rpc.dll`
Asegúrate de que `discord-rpc.dll` esté en el mismo directorio que el ejecutable del programa. Si no está presente, el programa mostrará un mensaje de error y no podrá iniciarse.

### 3. Compilación y Ejecución
1. Abre **Visual Basic 6** y carga el código.
2. Asegúrate de que la referencia a `discord-rpc.dll` esté en la carpeta del proyecto.
3. Ejecuta el código desde VB6 o compílalo en un ejecutable.
4. Al iniciar la aplicación, si Discord está abierto, verás el estado "Jugando AOForever" en tu perfil de Discord. (Cambien el nombre del juego por el suyo)

---

## Flujo del Programa
1. **Carga del formulario (`Form_Load`)**
   - Llama a `Ta()`, que inicia la conexión con Discord y configura el estado.

2. **Inicialización (`Ta`)**
   - Verifica la presencia de `discord-rpc.dll`.
   - Llama a `Discord_Initialize` para conectar con Discord.
   - Asigna memoria para almacenar los textos de estado.
   - Llama a `Discord_UpdatePresence` para actualizar la actividad del usuario en Discord.
   - Inicia un temporizador (`Timer1`) que ejecuta `Discord_RunCallbacks` cada segundo para mantener la conexión activa.

3. **Mantenimiento (`Timer1_Timer`)**
   - Ejecuta `Discord_RunCallbacks` para procesar los eventos de Discord y evitar desconexiones.

4. **Cierre del Programa (`Form_Unload`)**
   - Detiene el temporizador y libera la memoria asignada.
   - Llama a `Discord_Shutdown` para cerrar la conexión con Discord.

---

## Notas Importantes
- Si Discord no está abierto, la integración no funcionará.
- Este ejemplo solo establece un estado básico en Discord.

---

## Creditos
**Creador:** Kevin Korduner  
**Repositorio en GitHub:** https://github.com/KevinKorduner
**Integrado en:** AOForever

