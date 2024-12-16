library(shinydashboard)
library(DT)
library(forecast)
library(xts)
library(readxl)
library(ggplot2)
library(PerformanceAnalytics)
library(dplyr)
library(openxlsx) 
library(shinyWidgets)
library(shiny)

# Define la interfaz de usuario
ui <- fluidPage(
  titlePanel("Base de Datos"),
  tabsetPanel(
    tabPanel("Subir Datos",
             sidebarLayout(
               sidebarPanel(
                 fileInput("file_upload", "Subir archivo desde la computadora", accept = c(".csv", ".xlsx")),
                 br(),
                 selectInput("delimiter_select", "Selecciona el delimitador para CSV:",
                             choices = c("Coma (,)", "Punto y coma (;)", "Tabulación", "Otros")),
                 br(),
                 checkboxInput("header_checkbox", "¿El archivo tiene encabezados?", value = FALSE),
                 br(),
                 actionButton("load_data", "Cargar Datos"),
                 br(),
                 textOutput("message")  # Para mostrar mensajes de error o éxito
               ),
               mainPanel(
                 DTOutput("data_table")
               )
             )
    ),
    
    tabPanel("Análisis de Gráficos",
             sidebarLayout(
               sidebarPanel(
                 selectInput("columna_a_analizar", "Selecciona la columna para analizar:", ""),
                 actionButton("graficar_button", "Graficar"),
               ),
               mainPanel(
                 plotOutput("grafico_acf"),
                 plotOutput("grafico_pacf"),
                 plotOutput("grafico_residuals")
               )
             )
    ),
    tabPanel("Proyecciones",
             sidebarLayout(
               sidebarPanel(
                 selectInput("columna_a_proyectar", "Selecciona la columna a proyectar:", ""),
                 numericInput("periodos_proyeccion", "Número de periodos a proyectar:", value = 12),
                 numericInput("frequencia", "Frecuencia:", value = 12),
                 numericInput("p_value", "Valor de p:", value = 1),
                 numericInput("d_value", "Valor de d:", value = 1),
                 numericInput("q_value", "Valor de q:", value = 1),
                 numericInput("P_value", "Valor de P:", value = 1),
                 numericInput("D_value", "Valor de D:", value = 1),
                 numericInput("Q_value", "Valor de Q:", value = 1),
                 actionButton("proyectar_button", "Proyectar"),
                 downloadButton("download_excel", "Descargar Proyección MA"),  # Nuevo botón de descarga,
                 downloadButton("download_excel2", "Descargar Proyección MP")  # Nuevo botón de descarga
               ),
               mainPanel(
                 plotOutput("grafico_datos"),
                 plotOutput("grafico_proyeccion_mp"),
                 plotOutput("grafico_proyeccion_ma"),
                 tableOutput("criterios")
               )
             )
    ),
    tabPanel("Análisis Avanzado",
             sidebarLayout(
               sidebarPanel(
                 selectInput("analisis_tipo", "Tipo de Análisis", 
                             choices = c("Seleccione un análisis", "Regresión Dinámica", "Regresión Armónica", "Regresión Combinada")),
                 selectInput("columna_dependiente", "Selecciona la columna dependiente:", ""),
                 selectInput("columna_independiente", "Selecciona la columna independiente:",""),
                 numericInput("K", "K:", value = 2),
                 textInput("vectorValues", "Ingrese los valores independientes(separado por comas)", ""),
                 actionButton("analisis_button", "Realizar Análisis"),
                 downloadButton("download_excel3", "Descargar Proyección Avanzada")  # Nuevo botón de descarga
               ),
               mainPanel(
                 plotOutput("analisis_plot"),
               )
             )
    )
  )
)
# Define el servidor
server <- function(input, output, session) {
  
  # Almacena los datos cargados
  data <- reactiveVal(NULL)
  
  observeEvent(input$load_data, {
    req(input$file_upload)
    # Intenta cargar los datos del archivo seleccionado
    tryCatch({
      file_path <- input$file_upload$datapath
      file_ext <- tools::file_ext(file_path)
      
      # Determina el delimitador seleccionado
      delimiter <- switch(input$delimiter_select,
                          "Coma (,)" = ",",
                          "Punto y coma (;)" = ";",
                          "Tabulación" = "\t",
                          "Otros" = ","
      )
      
      if (file_ext %in% c("csv", "txt", "tsv")) {
        # Leer el archivo CSV con el delimitador seleccionado y manejar los encabezados
        if(input$header_checkbox){
          # Si el usuario indica que hay encabezados, leer el archivo normalmente
          raw_data <- read.csv(file_path, stringsAsFactors = FALSE, header = TRUE, sep = delimiter)
        } else {
          # Si no hay encabezados, asignar nombres de columnas por defecto
          raw_data <- read.csv(file_path, stringsAsFactors = FALSE, header = FALSE, sep = delimiter)
          colnames(raw_data) <- paste0("Columna", seq_len(ncol(raw_data)))
        }
      } else if (file_ext %in% c("xlsx", "xls")) {
        # Si es un archivo Excel, leerlo con readxl::read_excel
        # Utilizar col_types para detectar automáticamente fechas
        if(input$header_checkbox){
          raw_data <- readxl::read_excel(file_path, col_types = "guess")
        } else {
          # Si no hay encabezados, asignar nombres de columnas por defecto
          raw_data <- readxl::read_excel(file_path, col_names = FALSE)
          colnames(raw_data) <- paste0("Columna", seq_len(ncol(raw_data)))
        }
      } else {
        stop("Formato de archivo no admitido. Por favor, selecciona un archivo CSV o Excel.")
      }
      
      # Convertir todas las columnas que contienen fechas a formato YY-MM-DD
      date_columns <- sapply(raw_data, function(col) inherits(col, "POSIXt") || inherits(col, "Date"))
      raw_data[, date_columns] <- lapply(raw_data[, date_columns], function(x) {
        if (is.character(x)) {
          format(as.Date(x, tryFormats = c("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y")), "%Y-%m-%d")
        } else {
          x
        }
      })
      
      # Actualizar datos
      data(raw_data)
      output$message <- renderText("Datos cargados exitosamente.")
      
      # Actualizar las opciones de las columnas en el panel de proyecciones
      updateSelectInput(session, "columna_a_analizar", choices = colnames(raw_data))
      updateSelectInput(session, "columna_a_proyectar", choices = colnames(raw_data))
      updateSelectInput(session, "columna_dependiente", choices = colnames(raw_data))
      updateSelectInput(session, "columna_independiente", choices = colnames(raw_data))
    }, error = function(e) {
      output$message <- renderText(paste("Error al cargar los datos:", e$message))
    })
  })
  
  # Función para analizar graficos
  observeEvent(input$graficar_button, {
    req(!is.null(data()))
    
    # Obtener columnas seleccionadas
    
    columna_a_analizar <- input$columna_a_analizar
    frecuencia <- input$frequencia
    ts_data <- ts(data()[, columna_a_analizar], frequency = frecuencia)
    ts_data <- na.approx(ts_data)
    dife <- diff(ts_data)
    ACF <- acf(na.omit(dife))
    PACF <- pacf(na.omit(dife))
    
    modelo <- auto.arima(ts_data, lambda = BoxCox.lambda(ts_data))
    checkresiduals(modelo) 
    # Graficar
    output$grafico_acf <- renderPlot({
      plot(ACF, main = "ACF")
    })
    output$grafico_pacf <- renderPlot({
      plot(PACF, main = "PACF")
    })
    output$grafico_residuals <- renderPlot({
      checkresiduals(modelo) 
    })
  })
  
  observeEvent(input$proyectar_button, {
    req(!is.null(data()))
    
    # Obtener columnas seleccionadas
    columna_a_proyectar <- input$columna_a_proyectar
    Periodos_proyeccion <- input$periodos_proyeccion
    p <- input$p_value
    d <- input$d_value
    q <- input$q_value
    P <- input$P_value
    D <- input$D_value
    Q <- input$Q_value
    
    ts_data <- data()[, columna_a_proyectar]
    if (any(is.na(ts_data))) {
      output$message <- renderText("Error: Missing values present in the selected time series.")
      return()
    }
    frecuencia <- input$frequencia
    ts_data <- ts(data()[, columna_a_proyectar], frequency = frecuencia)
    
    
    # Fit ARIMA models
    modelo_arima_mp <- arima(ts_data, order = c(p, d, q), seasonal = list(order = c(P, D, Q)))
    modelo_arima_ma <- auto.arima(ts_data)
    Proyeccion_MP <- forecast(modelo_arima_mp, h = Periodos_proyeccion)
    Proyeccion_MA <- forecast(modelo_arima_ma, h = Periodos_proyeccion)
    
    # Cálculo de AIC y BIC en los modelos ARIMA
    AIC_MA <- AIC(modelo_arima_ma)
    AIC_MP <- AIC(modelo_arima_mp)
    BIC_MA <- BIC(modelo_arima_ma)
    BIC_MP <- BIC(modelo_arima_mp)
    # Graficar
    output$grafico_datos <- renderPlot({
      plot(ts(ts_data), main = "Serie Temporal", xlab = "Fecha", ylab = "Valor")
    })
    output$criterios <- renderTable({
      data.frame(
        Criterio = c("AIC Modelo Automático", "AIC Modelo Propio", "BIC Modelo Automático", "BIC Modelo Propio"),
        Valor = c(AIC_MA, AIC_MP, BIC_MA, BIC_MP)
      )
    }, rownames = TRUE)
    output$grafico_proyeccion_mp <- renderPlot({
      tryCatch({
        plot(Proyeccion_MP, main = "Proyección (Modelo Propio)", xlab = "Fecha", ylab = "Valor")
      }, error = function(e) {
        output$message <- renderText(paste("Error al graficar la proyección (MP):", e$message))
        return()
      })
    })
    
    output$grafico_proyeccion_ma <- renderPlot({
      tryCatch({
        plot(Proyeccion_MA, main = "Proyección (Modelo Automático)", xlab = "Fecha", ylab = "Valor")
      }, error = function(e) {
        output$message <- renderText(paste("Error al graficar la proyección (MA):", e$message))
        return()
      })
    })
    
  })
  
  observeEvent(input$analisis_button, {
    req(data()) # Asegura que los datos estén cargados
    
    columna_dependiente <- input$columna_dependiente
    columna_independiente <- input$columna_independiente
    Periodos_proyeccion <- input$periodos_proyeccion
    frecuencia <- input$frequencia
    k <- input$K
    
    x_future <- eventReactive(input$vectorValues, {
      as.numeric(unlist(strsplit(input$vectorValues, ",")))
    })
    y <- ts(data()[, columna_dependiente], frequency = frecuencia)
    y <- na.approx(y)
    x <- ts(data()[, columna_independiente], frequency = frecuencia)
    x <- na.approx(x)
    
    if(input$analisis_tipo == "Regresión Dinámica") {
      
      fit <- auto.arima(y, xreg = x)
      forecast_values <- forecast(fit, xreg = x_future())
      
    } else if(input$analisis_tipo == "Regresión Armónica") {
      
      fit2 <- auto.arima(y, xreg = fourier(x, K = k), seasonal = FALSE, lambda = 0)
      forecast_values <- forecast(fit2, xreg = fourier(x, K = k, h = Periodos_proyeccion))
      
    } else if(input$analisis_tipo == "Regresión Combinada"){
      
      
      arima_fit <- auto.arima(y)
      dynamic_fit <- auto.arima(y, xreg=x)
      harmonic_fit <- auto.arima(y,
                                 xreg = fourier(x, K = k),
                                 seasonal = FALSE, lambda = 0)
      
      arima_forecast <- forecast(arima_fit, h = Periodos_proyeccion)
      dynamic_forecast <- forecast(dynamic_fit, xreg = x_future(), h = Periodos_proyeccion)
      harmonic_forecast <- forecast(harmonic_fit, xreg = fourier(x, K = k, h = Periodos_proyeccion))
      
      forecast_values <- (arima_forecast$mean + dynamic_forecast$mean + harmonic_forecast$mean) / 3
      
      
      
    } else {
      output$analisis_text <- renderText("Por favor, selecciona un tipo de análisis.")
    }
    
    output$analisis_plot <- renderPlot({
      autoplot(forecast_values, xlab = "Tiempo", ylab = "Valor")
    })
    
  })
  
  # Muestra los datos en una tabla con formato de fecha YY-MM-DD
  output$data_table <- renderDT({
    if (!is.null(data())) {
      # Mostrar los datos en la tabla con formato YY-MM-DD
      datatable(data(), options = list(scrollX = TRUE))  # Add scrollX for horizontal scrolling
    }
  })
  
  # Función para guardar la proyección en un archivo de Excel y descargarlo MODELO AUTOMATICO
  output$download_excel <- downloadHandler(
    filename = function() {
      paste("ProyeccionMA_", input$columna_a_proyectar, ".xlsx", sep = "")
    },
    content = function(file) {
      if (!is.null(data())) {
        columna_a_proyectar <- input$columna_a_proyectar
        Periodos_proyeccion <- input$periodos_proyeccion
        p <- input$p_value
        d <- input$d_value
        q <- input$q_value
        P <- input$P_value
        D <- input$D_value
        Q <- input$Q_value
        
        ts_data <- data()[, columna_a_proyectar]
        
        Proyeccion_MA <- forecast(auto.arima(ts_data), h = Periodos_proyeccion)
        
        # Crear un nuevo libro de Excel
        wb <- createWorkbook()
        addWorksheet(wb, "Proyeccion")
        writeData(wb, "Proyeccion", as.data.frame(Proyeccion_MA), startCol = 1, startRow = 1)
        saveWorkbook(wb, file)
      }
    }
  )
  # Función para guardar la proyección en un archivo de Excel y descargarlo MODELO PROPIO
  output$download_excel2 <- downloadHandler(
    filename = function() {
      paste("ProyeccionMP_", input$columna_a_proyectar, ".xlsx", sep = "")
    },
    content = function(file) {
      if (!is.null(data())) {
        columna_a_proyectar <- input$columna_a_proyectar
        Periodos_proyeccion <- input$periodos_proyeccion
        p <- input$p_value
        d <- input$d_value
        q <- input$q_value
        P <- input$P_value
        D <- input$D_value
        Q <- input$Q_value
        
        ts_data <- data()[, columna_a_proyectar]
        
        Proyeccion_MP <- forecast(arima(ts_data, order = c(p, d, q), seasonal = list(order = c(P, D, Q))), h = Periodos_proyeccion)
        
        
        # Crear un nuevo libro de Excel
        wb <- createWorkbook()
        addWorksheet(wb, "Proyeccion")
        writeData(wb, "Proyeccion", as.data.frame(Proyeccion_MP), startCol = 1, startRow = 1)
        saveWorkbook(wb, file)
      }
    }
  ) 

  # Función para guardar la proyección avanzada en un archivo de Excel y descargarlo
  output$download_excel3 <- downloadHandler(
    filename = function() {
      "ProyeccionAvanzada.xlsx"
    },
    content = function(file) {
      if (!is.null(data())) {
        # Obtener los valores ingresados por el usuario
        columna_dependiente <- input$columna_dependiente
        columna_independiente <- input$columna_independiente
        Periodos_proyeccion <- input$periodos_proyeccion
        frecuencia <- input$frequencia
        k <- input$K
        
        x_future <- eventReactive(input$vectorValues, {
          as.numeric(unlist(strsplit(input$vectorValues, ",")))
        })
        y <- ts(data()[, columna_dependiente], frequency = frecuencia)
        y <- na.approx(y)
        x <- ts(data()[, columna_independiente], frequency = frecuencia)
        x <- na.approx(x)
        
        if(input$analisis_tipo == "Regresión Dinámica") {
          
          fit <- auto.arima(y, xreg = x)
          forecast_values <- forecast(fit, xreg = x_future())
          
        } else if(input$analisis_tipo == "Regresión Armónica") {
          
          fit2 <- auto.arima(y, xreg = fourier(x, K = k), seasonal = FALSE, lambda = 0)
          forecast_values <- forecast(fit2, xreg = fourier(x, K = k, h = Periodos_proyeccion))
          
        } else if(input$analisis_tipo == "Regresión Combinada"){
          
          
          arima_fit <- auto.arima(y)
          dynamic_fit <- auto.arima(y, xreg=x)
          harmonic_fit <- auto.arima(y,
                                     xreg = fourier(x, K = k),
                                     seasonal = FALSE, lambda = 0)
          
          arima_forecast <- forecast(arima_fit, h = Periodos_proyeccion)
          dynamic_forecast <- forecast(dynamic_fit, xreg = x_future(), h = Periodos_proyeccion)
          harmonic_forecast <- forecast(harmonic_fit, xreg = fourier(x, K = k, h = Periodos_proyeccion))
          
          forecast_values <- (arima_forecast$mean + dynamic_forecast$mean + harmonic_forecast$mean) / 3
          
          
          
        } else {
          output$analisis_text <- renderText("Por favor, selecciona un tipo de análisis.")
        }
        
        output$analisis_plot <- renderPlot({
          autoplot(forecast_values, xlab = "Tiempo", ylab = "Valor")
        })
        
     
        
        # Crear un nuevo libro de Excel y guardar los resultados de la proyección avanzada
        wb <- createWorkbook()
        addWorksheet(wb, "Proyeccion Dinamica")
        writeData(wb, "Proyeccion Dinamica", as.data.frame(forecast_values$mean), startCol = 1, startRow = 1)
        
        addWorksheet(wb, "Proyeccion Armonica")
        # Agregar los datos de la proyección armónica al archivo Excel
        writeData(wb, "Proyeccion Armonica", as.data.frame(forecast_values$mean), startCol = 1, startRow = 1)
        
        addWorksheet(wb, "Proyeccion Combinada")
        # Agregar los datos de la proyección combinada al archivo Excel
        writeData(wb, "Proyeccion Combinada", as.data.frame(forecast_values$mean), startCol = 1, startRow = 1)
        
        # Guardar el libro de Excel
        saveWorkbook(wb, file = file)
      }
    }
  )
}

# Crea la aplicación Shiny
shinyApp(ui, server)