# Evidencias reales — DAE Semana 11 (extra)

**Estudiante:** Edu Edward Galindo Huamán
**Repositorio:** https://github.com/edugalindo-jpg/dae-s11-hooks-del-agente
**Informe:** `Entregable_S11_Hooks_del_Agente_Edu_Galindo.pdf` (28 páginas)

Esta carpeta reúne las **16 capturas reales** y los **8 registros completos** que respaldan el
laboratorio. No hay montajes ni capturas retocadas: las imágenes de la aplicación salen del
servidor de desarrollo real y las de terminal se tomaron de la ventana de PowerShell.

## 1. Capturas de la aplicación (`capturas/01`–`10`)

| Archivo | Qué demuestra | Figura del informe |
|---|---|---|
| `01_carga_esqueleto.png` | Estado `loading`: esqueleto visible mientras la red simulada tarda 900 ms; el contador de la cabecera muestra «…». | Figura 7 |
| `02_planes_cargados.png` | Estado `ready`: los 6 planes reales (`P-01`…`P-06`) y el contador «Pendientes: 6». | Figura 8 |
| `03_error_red.png` | Estado `error` con `?fail=1`: aviso «No se pudieron cargar los planes» y botón «Reintentar». | Figura 9 |
| `04_busqueda_antigua.png` | `useDebounce` en espera: el campo ya muestra «values», pero la lista todavía no filtra. | Figura 10 |
| `05_busqueda_filtrada.png` | Pasados 300 ms, la lista se reduce a un solo plan (`P-03`): el valor «tranquilo» ya se aplicó. | Figura 11 |
| `06_contador_pendientes.png` | Tras aprobar un plan: «Pendientes: 5» en la cabecera y «Decididos: 1» en el panel. | Figura 12 |
| `07_ejecucion_r2_espera_aprobacion.png` | Un plan R2 (`P-01`) queda ensayado pero **no aplicado**: pide confirmación explícita. | Figura 13 |
| `08_ejecucion_aplicada.png` | Tras «Confirmar y aplicar»: estado `done` y hash real `d68ebfa913ec`. | Figura 14 |
| `09_persistencia_recarga.png` | Después de recargar: «Pendientes: 4» y «Decididos: 2»; las decisiones siguen en `localStorage`. | Figura 15 |
| `10_vista_movil.png` | Vista móvil (390 × 844): la bandeja, el buscador y las tarjetas se adaptan. | Figura 16 |

Las capturas se tomaron con un viewport de 900 × 900 a `deviceScaleFactor: 2` y se recortaron a
la columna real de la app (880 px). Antes de cada disparo, `capturar_app.py` comprueba el estado
visible; si el texto esperado no está en pantalla, no guarda la imagen. Por eso los PNG miden
1784 px de ancho y el texto se lee con nitidez.

## 2. Capturas de terminal (`capturas/11`–`15`)

| Archivo | Qué demuestra | Figura del informe |
|---|---|---|
| `11_terminal_test_antes.png` | `npm test` antes de corregir: 17 pasan y 40 fallan; código de salida 1. | Figura 2 |
| `12_terminal_test_ai_antes.png` | `npm run test:ai` antes: 37 pasan y 20 fallan; los seis fallos de la IA. | Figura 3 |
| `13_terminal_test_despues.png` | `npm test` después: 5 archivos y 57/57 pruebas en verde; código 0. | Figura 4 |
| `14_terminal_test_ai_despues.png` | `npm run test:ai` después: 57/57 en verde sobre `src/ai`. | Figura 5 |
| `15_terminal_build.png` | `npm run build`: compilación correcta, 44 módulos transformados. | Figura 6 |

Las capturas de terminal se hicieron con `PrintWindow` sobre la ventana de PowerShell, no con un
recorte de pantalla: así la evidencia sigue siendo válida aunque otra ventana tape la consola.

## 3. Estructura del proyecto (`capturas/16`)

| Archivo | Qué demuestra | Figura del informe |
|---|---|---|
| `16_vscode_estructura.png` | Árbol real del repositorio y `src/hooks/usePlans.js` abierto en VS Code. | Figura 1 |

## 4. Registros completos (`logs/`)

| Archivo | Contenido |
|---|---|
| `01_test_inicial.txt` | Salida completa de `npm test` antes de implementar: 17/57. |
| `02_test_ai_inicial.txt` | Salida completa de `npm run test:ai` antes de corregir: 37/57. |
| `03_hooks_dirigidos.txt` | Corrida dirigida a los hooks propios. |
| `04_test_final.txt` | `npm test` final: 57/57. |
| `05_test_ai_final.txt` | `npm run test:ai` final: 57/57. |
| `06_build.txt` | `npm run build` final. |
| `11_antes_test_consola.txt` | Consola original del «antes» de `npm test`. |
| `12_antes_test_ai_consola.txt` | Consola original del «antes» de `npm run test:ai`. |

## 5. Reproducibilidad (`scripts_reproduccion/`)

Se incluyen los scripts que generaron las capturas:

- `capturar_app.py`: recorre los estados de la app y verifica el texto visible antes de capturar.
- `capturar_ventana.ps1`: captura la ventana de PowerShell con `PrintWindow`.
- `capturar_vscode.ps1`: captura VS Code.
- `ocr.ps1`: comprobación OCR de que el texto clave está en los PNG.

Para repetirlas:

```bash
npm run dev        # deja la app en http://localhost:5173
python docs/scripts/capturar_app.py
```

## 6. Cobertura de los criterios

- **Tres hooks propios:** `capturas/04`–`06` y `07`–`09` muestran `useDebounce`, `usePlans` y `useAgentRun` en uso real.
- **Seis fallos de la IA corregidos:** `capturas/11`–`12` (antes) frente a `13`–`14` (después), con los registros `01`–`05`.
- **Propiedad del estado:** `02`, `06` y `09` prueban que el contador se calcula en `App` y que `Header` solo lo recibe; la justificación está en la sección 2.4 del informe.
- **Pruebas y build:** `13`, `14` y `15`, más los logs `04`–`06`.
- **Evidencia real y reproducible:** scripts incluidos y comandos del informe (Anexo C).

## 7. Integridad

`SHA256SUMS.txt` contiene el hash SHA-256 de cada archivo de esta carpeta. Para comprobarlo:

```powershell
Get-FileHash .\capturas\*.png -Algorithm SHA256
```

Los archivos también están versionados en el repositorio, en `docs/capturas/` y `docs/logs/`.
