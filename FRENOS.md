# Frenos de gasto

## Límites operativos
| Regla | Valor |
|---|---|
| Obreras simultáneas | máx. 3 |
| Obreras nuevas por pedido sin confirmar | máx. 3 |
| Reintentos de una obrera caída | 1 automático; después se consulta al usuario |

## Freno por uso del plan
- La reina **no puede leer** el % de uso del plan. Ese dato lo ve el usuario en la configuración de su cuenta.
- Cuando el usuario diga **"FRENO"** (por ejemplo, porque llegó al 80 %), la reina envía FRENO a todas las obreras activas.
- Ante FRENO, cada obrera deja un punto de control en su bitácora, hace push y se detiene. Nadie corta en seco.
- Al recuperar cupo, el usuario dice "retoma" y cada obrera sigue desde su punto de control.

## Ahorro de tokens
- Un ajuste puntual de un solo encargo se hace directamente en la sesión de esa obrera, no a través de la reina (si pasa por la reina, se paga dos veces).
- A la reina se le habla para: arrancar varios trabajos, cambios que afectan a varios encargos, decisiones y reanudaciones.
