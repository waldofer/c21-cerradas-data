# Reporte de corrección de precios históricos

Dataset: 5159 -> 5143 registros.
Candidatos procesados: 482.

## Resumen por causa

| Causa | Acción | Casos |
|---|---|---|
| validado | corregir | 401 |
| ambiguo_revision_manual | marcar | 26 |
| no_es_precio_de_venta | eliminar | 16 |
| zona_sin_comparables | corregir | 15 |
| digito_faltante | corregir | 11 |
| superficie_corrupta | corregir_y_marcar | 6 |
| parcela_grande | corregir | 6 |
| sin_pares | marcar | 1 |

## Pendientes de revisión manual

Quedan marcados con `px` y **sin corregir**.

| ID | Fecha | Subtipo | Barrio | pr (Gs) | m² | Detalle |
|---|---|---|---|---|---|---|
| 12657 | 2023-04-14 | terreno | Altos | 85,000,000 | 1080.0 | factor 0.178 con 5 pares sólidos (L1). |
| 12656 | 2023-04-14 | terreno | Altos | 85,000,000 | 1080.0 | factor 0.178 con 5 pares sólidos (L1). |
| 3360 | 2023-04-14 | terreno | Altos | 80,000,000 | 1080.0 | factor 0.168 con 5 pares sólidos (L1). |
| 12031 | 2023-04-19 | terreno | Altos | 35,000,000 | 551.97 | factor 0.144 con 25 pares sólidos (L1). |
| 11797 | 2023-04-19 | terreno | Altos | 52,500,000 | 777.84 | factor 0.169 con 13 pares sólidos (L1). |
| 6671 | 2023-05-16 | casa | Altos | 380,000,000 | 1312.0 | factor 0.060 con 4 pares sólidos (L2). |
| 13084 | 2023-05-12 | terreno | Emboscada | 40,000,000 | 2533.0 | factor 0.080 con 3 pares sólidos (L2). |
| 13083 | 2023-05-12 | terreno | Emboscada | 40,000,000 | 2533.0 | factor 0.080 con 3 pares sólidos (L2). |
| 13082 | 2023-05-12 | terreno | Emboscada | 40,000,000 | 2533.0 | factor 0.080 con 3 pares sólidos (L2). |
| 13081 | 2023-05-12 | terreno | Emboscada | 40,000,000 | 2533.0 | factor 0.080 con 3 pares sólidos (L2). |
| 13080 | 2023-05-12 | terreno | Emboscada | 40,000,000 | 2533.0 | factor 0.080 con 3 pares sólidos (L2). |
| 7051 | 2023-05-12 | terreno | Emboscada | 40,000,000 | 2533.0 | factor 0.080 con 3 pares sólidos (L2). |
| 1551 | 2020-12-18 | terreno | Luque | 115,000,000 | 360.67 | factor 0.188 con 4 pares sólidos (L1). |
| 13696 | 2023-06-08 | terreno | Luque | 180,000,000 | 1000.0 | factor 0.186 con 25 pares sólidos (L1). |
| 1973 | 2021-11-17 | terreno | Mariano Roque Alonso | 85,000,000 | 360.0 | factor 0.085 con 4 pares sólidos (L1). |
| 1971 | 2021-11-17 | terreno | Mariano Roque Alonso | 85,000,000 | 360.0 | factor 0.085 con 4 pares sólidos (L1). |
| 3915 | 2021-11-23 | terreno | Mariano Roque Alonso | 130,000,000 | 360.0 | factor 0.130 con 4 pares sólidos (L1). |
| 1737 | 2020-10-09 | terreno | San Bernardino | 80,000,000 | 704.0 | factor 0.170 con 72 pares sólidos (L1). |
| 1381 | 2020-10-23 | terreno | San Bernardino | 85,000,000 | 673.0 | factor 0.188 con 72 pares sólidos (L1). |
| 1655 | 2020-12-22 | terreno | San Bernardino | 420,000,000 | 4217.0 | factor 0.129 con 5 pares sólidos (L1). |
| 2876 | 2021-02-01 | terreno | San Bernardino | 210,000,000 | 50.0 | factor 6.806 con 169 pares sólidos (L2). |
| 5181 | 2022-02-21 | terreno | San Bernardino | 95,000,000 | 787.5 | factor 0.198 con 265 pares sólidos (L1). |
| 9314 | 2022-11-16 | terreno | San Bernardino | 125,000,000 | 1680.0 | factor 0.151 con 62 pares sólidos (L1). PISTA: prL*6100 da 4.4x el pr, el monto puede estar mal cargado. |
| 6293 | 2022-11-16 | terreno | San Bernardino | 125,000,000 | 1680.0 | factor 0.151 con 62 pares sólidos (L1). PISTA: prL*6100 da 4.4x el pr, el monto puede estar mal cargado. |
| 3510 | 2022-05-31 | terreno | Ypacaraí | 54,000,000 | 384.0 | factor 0.169 con 21 pares sólidos (L1). |
| 3512 | 2022-06-25 | terreno | Ypacaraí | 55,000,000 | 386.25 | factor 0.171 con 21 pares sólidos (L1). |

## Eliminados del dataset de ventas

| ID | Fecha | Subtipo | Barrio | Detalle |
|---|---|---|---|---|
| 1539 | 2020-07-24 | nave | Barrio Jara | 7,700,000 Gs -> USD 1,110: no es precio de venta |
| 2035 | 2020-12-15 | terreno | San Juan del Paraná | 20,000,000 Gs -> USD 2,861: no es precio de venta |
| 2378 | 2021-01-26 | casa | San Lorenzo | 6,500,000 Gs -> USD 934: no es precio de venta |
| 3049 | 2021-04-12 | propiedad-agricola | Ybycuí | 10,000,000 Gs -> USD 1,602: no es precio de venta |
| 3300 | 2021-06-07 | terreno | Pirebebuy | 12,500,000 Gs -> USD 1,862: no es precio de venta |
| 3299 | 2021-06-07 | terreno | Pirebebuy | 12,500,000 Gs -> USD 1,862: no es precio de venta |
| 3298 | 2021-06-07 | terreno | Pirebebuy | 15,000,000 Gs -> USD 2,235: no es precio de venta |
| 4416 | 2021-09-10 | departamento | Luque, Central | 14,755,000 Gs -> USD 2,136: no es precio de venta |
| 3169 | 2021-12-20 | casa-duplex | San Bernardino | 10,000,000 Gs -> USD 1,470: no es precio de venta |
| 6679 | 2022-05-07 | casa | Los Laureles | 6,860,000 Gs -> USD 1,002: no es precio de venta |
| 8181 | 2022-09-16 | casa-en-condominio | Las Lomas | 9,000,000 Gs -> USD 1,289: no es precio de venta |
| 9147 | 2022-09-30 | local | Barrio Jara | 7,000,000 Gs -> USD 987: no es precio de venta |
| 6172 | 2022-11-10 | departamento | Las Mercedes | 5,800,000 Gs -> USD 822: no es precio de venta |
| 9567 | 2023-01-17 | departamento | Ycuá Satí | 6,000,000 Gs -> USD 808: no es precio de venta |
| 11647 | 2023-03-06 | nave | San Vicente | 6,300,000 Gs -> USD 877: no es precio de venta |
| 14246 | 2023-06-29 | casa-duplex | Los Laureles | 6,000,000 Gs -> USD 825: no es precio de venta |

## digito_faltante (11)

| ID | Fecha | Barrio | Detalle |
|---|---|---|---|
| 538 | 2020-02-26 | Itá Enramada, Asunción | pr 70,000,000 -> 700,000,000 Gs (dígito restituido), tasa 6,514.83 |
| 13094 | 2023-06-05 | Emboscada | pr 28,000,000 -> 280,000,000 Gs (dígito restituido), tasa 7,245.38 |
| 13093 | 2023-06-05 | Emboscada | pr 28,000,000 -> 280,000,000 Gs (dígito restituido), tasa 7,245.38 |
| 13092 | 2023-06-05 | Emboscada | pr 28,000,000 -> 280,000,000 Gs (dígito restituido), tasa 7,245.38 |
| 13091 | 2023-06-05 | Emboscada | pr 28,000,000 -> 280,000,000 Gs (dígito restituido), tasa 7,245.38 |
| 13090 | 2023-06-05 | Emboscada | pr 28,000,000 -> 280,000,000 Gs (dígito restituido), tasa 7,245.38 |
| 13089 | 2023-06-05 | Emboscada | pr 28,000,000 -> 280,000,000 Gs (dígito restituido), tasa 7,245.38 |
| 13088 | 2023-06-05 | Emboscada | pr 28,000,000 -> 280,000,000 Gs (dígito restituido), tasa 7,245.38 |
| 13086 | 2023-06-05 | Emboscada | pr 28,000,000 -> 280,000,000 Gs (dígito restituido), tasa 7,245.38 |
| 13085 | 2023-06-05 | Emboscada | pr 28,000,000 -> 280,000,000 Gs (dígito restituido), tasa 7,245.38 |
| 7054 | 2023-06-05 | Emboscada | pr 28,000,000 -> 280,000,000 Gs (dígito restituido), tasa 7,245.38 |

## superficie_corrupta (6)

| ID | Fecha | Barrio | Detalle |
|---|---|---|---|
| 1100 | 2020-02-17 | Itauguá | Precio OK (USD 33,762) pero superficie 2.0 m² es imposible |
| 5116 | 2021-12-09 | San Bernardino | Precio OK (USD 13,907) pero superficie 1.0 m² es imposible |
| 11487 | 2023-03-03 | Luque | Precio OK (USD 208,789) pero superficie 1.0 m² es imposible |
| 12329 | 2023-03-15 | Loma Pytá | Precio OK (USD 76,218) pero superficie 1.0 m² es imposible |
| 11957 | 2023-05-31 | Caacupé | Precio OK (USD 186,952) pero superficie 17.7 m² es imposible |
| 33045 | 2025-07-23 | Julián Augusto Saldívar | Precio OK (USD 12,002) pero superficie 1.14 m² es imposible |

## sin_pares (1)

| ID | Fecha | Barrio | Detalle |
|---|---|---|---|
| 822 | 2020-05-16 | Villa Elisa | No hay comparables suficientes (subtipo bodega) |
