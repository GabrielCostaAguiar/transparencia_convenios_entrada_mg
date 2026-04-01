library(dplyr)   # , purrr, tibble, etc.
library(readxl)
library(lubridate)   # year(Sys.Date())
library(googledrive)
library(readr)
library(openxlsx)
stwd <- "C:\\Users\\M1578465\\Documents\\DCGCE\\Automacoes\\Transparência SIGCON"

#Passo 1: Autenticar o drive pra conseguir ler o link dos arquivos

drive_auth_configure(path = "C:\\Users\\M1578465\\Desktop\\API TEs\\json.json")  # arquivo do token
drive_auth(email = "dcgce.seplag@gmail.com")
drive_user()

#Passo 2: Ler o link de cada arquivo do drive

lista_instrumentos <- drive_ls("transparencia") #Lendo arquivos na pasta

lista_instrumentos1 <- lista_instrumentos %>% 
  mutate(instrumento_drive = paste0("https://drive.google.com/file/d/", id, "/view")) %>% 
  mutate("Doc_autorizativo" = substr(name, 1,nchar(lista_instrumentos$name) - 4))


#Passo 3: Ler base de dados e criar coluna para os links
controle_sei <- read.xlsx("C:\\Users\\M1578465\\Documents\\DCGCE\\Automacoes\\Transparência SIGCON\\Controle SEI.xlsx")
controle_sei1 <- controle_sei %>% 
  select("Nº.SIAFI_(SIGCON)", "Instrumento") %>% 
  rename("Código.SIAFI" = "Nº.SIAFI_(SIGCON)") %>% 
  rename("Doc_autorizativo" = "Instrumento") %>% 
  mutate(Código.SIAFI = as.character(Código.SIAFI))


consultas_sigcon <- read.xlsx("C:\\Users\\M1578465\\Documents\\DCGCE\\Automacoes\\Transparência SIGCON\\Consultas SIGCON - Registros Instrumentos - ATUALIZADO.xlsx")

consultas_sigcon1 <- consultas_sigcon %>% 
  mutate(Código.SIAFI = as.character(Código.SIAFI))

consultas_sigcon2 <- left_join(consultas_sigcon1, controle_sei1, by = "Código.SIAFI")

colnames(consultas_sigcon_final)

#Adicionar o link do drive
consultas_sigcon_final <- left_join(consultas_sigcon2, lista_instrumentos1, by = "Doc_autorizativo") %>% 
  select(-32, -33, -34)

consultas_sigcon_final1 <- consultas_sigcon_final %>% 
  mutate(`Inteiro.teor.do.Instrumento.-.Sigcon` = if_else(
    is.na(`Inteiro.teor.do.Instrumento.-.Sigcon`) & !is.na(instrumento_drive),
    instrumento_drive,
    `Inteiro.teor.do.Instrumento.-.Sigcon`
  ))
  
consultas_sigcon_final2 <- consultas_sigcon_final1 %>% 
  mutate(Data.Publicação = as.Date(Data.Publicação, format = "%d/%m/%Y")) %>% 
  filter(
    Data.Publicação >= as.Date("01/01/2022", format = "%d/%m/%Y"),          # Data comparável
    Situação != "BLOQUEADO",                           # Remove bloqueados
    !(is.na(`Inteiro.teor.do.Instrumento.-.Sigcon`) &  # Mantém se pelo menos uma tem valor
        is.na(`Inteiro.Teor.do.Instrumento.-.TransfereGov`))) %>% 
  select(-instrumento_drive) %>% 
  mutate(Data.Publicação = format(Data.Publicação, "%d/%m/%Y"))

# Criando um workbook
wb <- createWorkbook()

# Adicionando uma aba
addWorksheet(wb, "Consulta_SIGCON")

# Criando um estilo para o cabeçalho (negrito)
header_style <- createStyle(
  fontName = "Tahoma",
  fontSize = 8,
  fontColour = "black",
  textDecoration = "bold", # Negrito
  halign = "center",       # Alinhamento horizontal centralizado
  valign = "center",       # Alinhamento vertical centralizado
  border = "TopBottomLeftRight",       # Borda em todos os lados
  borderColour = "black"
)
# Criando um estilo para os dados (sem negrito)
data_style <- createStyle(
  fontName = "Tahoma",
  fontSize = 8,
  fontColour = "black",
  halign = "left",  # Alinhamento horizontal à esquerda
  valign = "center", # Alinhamento vertical centralizado
  border = "TopBottomLeftRight",       # Borda em todos os lados
  borderColour = "black"
)

# Escrevendo o DataFrame na aba
writeData(wb, sheet = "Consulta_SIGCON", x = consultas_sigcon_final2)

# Aplicando estilo ao cabeçalho (primeira linha)
addStyle(wb, sheet = "Consulta_SIGCON", style = header_style, rows = 1, cols = 1:(ncol(consultas_sigcon_final2) + 1), gridExpand = TRUE)

# Aplicando estilo aos dados (linhas restantes)
addStyle(wb, sheet = "Consulta_SIGCON", style = data_style, rows = 2:(nrow(consultas_sigcon_final2) + 1), cols = 1:(ncol(consultas_sigcon_final2) + 1), gridExpand = TRUE)

# Ajustando largura das colunas automaticamente
setColWidths(wb, sheet = "Consulta_SIGCON", cols = 1:(ncol(consultas_sigcon_final2) + 1), widths = 15)


#Cria o arquivo atualizado
nome_arquivo <- paste("Consultas SIGCON - Instrumentos 2022 a 2025 - ATUALIZADO", format(Sys.Date(), "%d-%m-%Y"), ".xlsx", sep = " ")
saveWorkbook(wb, nome_arquivo, overwrite = TRUE)

#Subindo documento no drive
#drive_auth(path = "C:\\Users\\M1578465\\Documents\\DCGCE\\Automacoes\\Texto decreto\\service_account.json")  # arquivo do token
#drive_user()

# Faz upload para o Google Drive
#drive_upload(
 # media = "C:\\Users\\M1578465\\Documents\\DCGCE\\Automacoes\\Transparência SIGCON\\PDFs",
  #path = as_id("1zjADlxXp4NqF3pwd6Eez0ciofFQDN3jY"),  # ou apenas o nome da pasta
 # name = "Instrumentos",
 # overwrite = TRUE   # substitui se já existir
#)

#print("Upload feito com sucesso!")

