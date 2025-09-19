#######################################################
################ RECOLECCIÓN DE DATOS     #############
################ MODELOS VaR              #############
#######################################################
cat("\014") # Limpiar la consola
while(dev.cur() > 1) dev.off() # Limpiar todos los gráficos
rm(list = ls()) # Limpiar el entorno global

library(dplyr)
library(openxlsx)
library(lubridate)
library(extrafont) # New fonts
library(janitor)
library(reshape2)
library(zoo)
library(rugarch)
# Cargar las fuentes en el dispositivo gráfico
#loadfonts(device = "win")


ruta <- "Z:/03_Investigaciones_Economicas/3. Modelos/1. Forecast/VaR/"
libro <- "Value at Risk Colombia.xlsx"
hoja  <- "Data"

# Llamar la hoja de excel principal donde tengo los datos, pero sin los nombres de las columnas
df <- read.xlsx(paste0(ruta, libro),sheet = hoja, startRow = 3, colNames = T)
glimpse(df)

# Estos son los nombres de las columnas
df_names <- read.xlsx(paste0(ruta, libro), sheet = hoja, startRow = 2, colNames = F)
glimpse(df_names)

## filtramos a solo la primera fila. equivale a la fila 2 de excel
df_names <- df_names[1,] %>% as.character()

colnames(df) <- df_names
glimpse(df)

if (is.numeric(df[,1])) {
  # SI EL FORMATO ESTA EN EL FORMATO EXCEL 45170, 45173, ETC
  df[,1] <- as.Date(df[,1], origin = "1899-12-30")
} else {
  # SI YA ESTA EN FORMATO DTTM, RELAX, ESTA OK. SOLO CONVERTIRLO A DATE PARA LUEGO USAR GGPLOT2
  df[,1] <- as.Date(df[,1])
}
# CONFIRMAMOS QUE SE HAYA CAMBIADO
glimpse(df)

df <- df %>% clean_names()

# AHORA AGREGAR LOS ARCHIVOS DE NDF
# Buscar automáticamente el archivo más reciente que comience con "Liquidaciones_NDF"
archivos <- list.files(path = ruta, pattern = "^Liquidaciones_NDF.*\\.xlsx$", full.names = TRUE)

# Validar si se encontró al menos uno
if (length(archivos) == 0) {
  stop("No se encontró ningún archivo que comience con 'Liquidaciones_NDF' en la ruta especificada.")
}

# Seleccionar el más reciente por fecha de modificación
archivo_reciente <- archivos[which.max(file.info(archivos)$mtime)]
archivo_reciente

# Cargar la hoja de los NDFs
df_ndf <- read.xlsx(archivo_reciente, sheet = "DATOS", colNames = TRUE) %>%
  clean_names() # siempre ponerlo en minuscula :)
glimpse(df_ndf)


# Convertir fechas
cols_fecha <- c("fecdat", "desde", "hasta", "f_valor", "f_venci")
df_ndf[cols_fecha] <- lapply(df_ndf[cols_fecha], function(x) as.Date(x, origin = "1899-12-30"))
glimpse(df_ndf)

# Colocar la fecha de pago (solo algunas veces hay t+2, pero es muy puntual)
df_ndf$fecha_pago <- df_ndf$f_venci + 1

df_ndf <- df_ndf %>% 
  mutate(
    year = lubridate::year(f_venci)
    , month = lubridate::month(f_venci)
  )

unique(df_ndf$moneda) # validar cuántas monedas hay
unique(df_ndf$clase_ope) # validar clases de operaciones, solo deben existir dos


# Colocar los tipos de cambio vigentes

usdcop = 3897.57 # TRM de la Superfinanciera
eurusd = 1.1807 # referencia del ultimo dato que reports BCE
eurcop = usdcop * eurusd # simplemente multiplicar
date_today = lubridate::today()



# Creamos las cuatro bases para no hacer tantos condicionales

# USD y compra
df_usd_c <- df_ndf %>% filter(moneda == "USD", clase_ope == "C")

df_usd_c <- df_usd_c %>%
  mutate(
    liquidacion_ndf_sensibilizada = nominal * (usdcop - cambio_fwd), # formula para liquidar ndfs
    liquidacion_ndf_final = case_when(
      fecha_pago < date_today ~ liquidacion_ndf, # columna original
      TRUE ~ liquidacion_ndf_sensibilizada # columna calculada arriba
    )
  )

# valor nominal de USD compra
nominal_usd_c <- sum(df_usd_c$liquidacion_ndf_final)
nominal_usd_c

# USD y venta
df_usd_v <- df_ndf %>% filter(moneda == "USD", clase_ope == "V")

df_usd_v <- df_usd_v %>%
  mutate(
    liquidacion_ndf_sensibilizada = (nominal*-1)* (usdcop - cambio_fwd), # formula para liquidar ndfs
    liquidacion_ndf_final = case_when(
      fecha_pago < date_today ~ liquidacion_ndf, # columna original
      TRUE ~ liquidacion_ndf_sensibilizada # columna calculada arriba
    )
  )

# valor nominal de USD venta
nominal_usd_v <- sum(df_usd_v$liquidacion_ndf_final)
nominal_usd_v

# EUR y compra
df_eur_c <- df_ndf %>% filter(moneda == "EUR", clase_ope == "C")

df_eur_c <- df_eur_c %>%
  mutate(
    liquidacion_ndf_sensibilizada = (nominal)* (eurusd - cambio_fwd)*usdcop, # formula para liquidar ndfs
    liquidacion_ndf_final = case_when(
      fecha_pago < date_today ~ liquidacion_ndf, # columna original
      TRUE ~ liquidacion_ndf_sensibilizada # columna calculada arriba
    )
  )

# valor nominal de USD compra
nominal_eur_c <- sum(df_eur_c$liquidacion_ndf_final)
nominal_eur_c

# EUR y venta
df_eur_v <- df_ndf %>% filter(moneda == "EUR", clase_ope == "V")

df_eur_v <- df_eur_v %>%
  mutate(
    liquidacion_ndf_sensibilizada = (nominal*-1)* (eurusd - cambio_fwd)*usdcop, # formula para liquidar ndfs
    liquidacion_ndf_final = case_when(
      fecha_pago < date_today ~ liquidacion_ndf, # columna original
      TRUE ~ liquidacion_ndf_sensibilizada # columna calculada arriba
    )
  )

# valor nominal de USD compra
nominal_eur_v <- sum(df_eur_v$liquidacion_ndf_final)
nominal_eur_v



# Escenario 1: VaR histórico simple sobre retornos de tipo de cambio
df <- df %>% arrange(date)
df$retornos_usdcop <- c(NA,diff(log(df$usdcop))*100)
df$retornos_eurusd <- c(NA,diff(log(df$eurusd))*100)
glimpse(df)

# VaR al 95% de confianza
var_historico_usdcop <- quantile(df$retornos_usdcop, probs = 0.05, na.rm = TRUE)
var_historico_usdcop

var_historico_eurusd <- quantile(df$retornos_eurusd, probs = 0.05, na.rm = TRUE)
var_historico_eurusd

# Escenario 2: VaR paramétrico usando media y desviación estándar
media_usdcop <- mean(df$retornos_usdcop, na.rm = TRUE)
sd_usdcop <- sd(df$retornos_usdcop, na.rm = TRUE)

# Supuesto de normalidad
var_parametrico_usdcop <- qnorm(0.05, mean = media_usdcop, sd = sd_usdcop)
var_parametrico_usdcop


media_eurusd <- mean(df$retornos_eurusd, na.rm = TRUE)
sd_eurusd <- sd(df$retornos_eurusd, na.rm = TRUE)

# Supuesto de normalidad
var_parametrico_eurusd <- qnorm(0.05, mean = media_eurusd, sd = sd_eurusd)
var_parametrico_eurusd


# Escenario 3: Simulación Monte Carlo
set.seed(123)
simulaciones_usdcop <- rnorm(1000000, mean = media_usdcop, sd = sd_usdcop)
var_montecarlo_usdcop <- quantile(simulaciones_usdcop, probs = 0.05)
var_montecarlo_usdcop


simulaciones_eurusd <- rnorm(1000000, mean = media_eurusd, sd = sd_eurusd)
var_montecarlo_eurusd <- quantile(simulaciones_eurusd, probs = 0.05)
var_montecarlo_eurusd

# Escenario 4: VaR con GARCH(1,1)
library(rugarch)

spec <- ugarchspec(variance.model = list(model = "sGARCH", garchOrder = c(1,1)),
                   mean.model = list(armaOrder = c(0,0)),
                   distribution.model = "norm")


retornos_usdcop <- na.omit(df$retornos_usdcop)



fit <- ugarchfit(spec, data = retornos_usdcop)
sim <- ugarchsim(fit, n.sim = 10000)
var_garch <- quantile(fitted(sim), probs = 0.05)
var_garch
