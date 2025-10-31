library(shiny)
library(plotly)
library(dplyr)
library(readxl)
library(lubridate)
library(zoo)
library(DT)
library(writexl)

# ----------------------
# UI 部分
# ----------------------
ui <- fluidPage(
  # 添加自定义CSS样式
  tags$head(
    tags$style(HTML("
      /* 整体页面样式 */
      body {
        font-family: 'Microsoft YaHei', Arial, sans-serif;
        font-size: 14px;
      }
      
      /* 标题样式 - 居中 */
      .title-panel {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 20px;
        border-radius: 8px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        text-align: center;
      }
      
      /* 自定义标题样式 */
      .custom-title {
        font-size: 28px;
        font-weight: bold;
        margin: 0;
        padding: 10px 0;
        text-align: center;
      }
      
      /* 侧边栏样式 */
      .sidebar-panel {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        height: fit-content;
        position: sticky;
        top: 20px;
      }
      
      /* 主面板样式 */
      .main-panel {
        background-color: #ffffff;
        border-radius: 8px;
        padding: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        min-height: 600px;
      }
      
      /* 选择输入框样式 */
      .select-input {
        width: 100% !important;
        margin-bottom: 15px;
      }
      
      /* 文件上传样式 */
      .file-input {
        width: 100% !important;
        margin-bottom: 15px;
      }
      
      /* 按钮样式 */
      .action-button {
        width: 100%;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 12px 20px;
        border-radius: 6px;
        font-size: 16px;
        font-weight: bold;
        margin: 10px 0;
        transition: all 0.3s ease;
      }
      
      .action-button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
      }
      
      /* 保存按钮样式 */
      .save-button {
        width: 100%;
        background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
        color: white;
        border: none;
        padding: 12px 20px;
        border-radius: 6px;
        font-size: 16px;
        font-weight: bold;
        margin: 10px 0;
        transition: all 0.3s ease;
      }
      
      .save-button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
      }
      
      /* 进度条容器 */
      .progress-container {
        margin: 15px 0;
        padding: 10px;
        background: #f8f9fa;
        border-radius: 6px;
        border: 1px solid #e9ecef;
      }
      
      /* 进度条样式 */
      .progress {
        height: 20px;
        margin-bottom: 5px;
        border-radius: 10px;
        background-color: #e9ecef;
        overflow: hidden;
      }
      
      .progress-bar {
        height: 100%;
        background: linear-gradient(90deg, #667eea, #764ba2);
        transition: width 0.3s ease;
        border-radius: 10px;
      }
      
      /* 进度文本 */
      .progress-text {
        font-size: 12px;
        color: #666;
        text-align: center;
        margin-top: 5px;
      }
      
      /* 水平分隔线 */
      .custom-hr {
        border: 0;
        height: 1px;
        background: linear-gradient(to right, transparent, #667eea, transparent);
        margin: 20px 0;
      }
      
      /* 标签页样式 */
      .nav-tabs {
        background-color: #f1f3f4;
        border-radius: 8px 8px 0 0;
        padding: 5px 5px 0 5px;
      }
      
      .nav-tabs > li > a {
        color: #555;
        font-weight: bold;
        border-radius: 6px 6px 0 0;
        margin-right: 2px;
      }
      
      .nav-tabs > li.active > a {
        background-color: #667eea;
        color: white;
        border: none;
      }
      
      /* 图表容器 */
      .plot-container {
        background: white;
        border-radius: 8px;
        padding: 15px;
        margin-bottom: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
      }
      
      /* 数据表容器 */
      .data-table-container {
        background: white;
        border-radius: 8px;
        padding: 15px;
        margin-top: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
      }
      
      /* 欢迎信息样式 */
      .welcome-message {
        text-align: center;
        padding: 60px 20px;
        color: #666;
      }
      
      .welcome-icon {
        font-size: 64px;
        color: #667eea;
        margin-bottom: 20px;
      }
      
      .welcome-title {
        font-size: 24px;
        font-weight: bold;
        margin-bottom: 15px;
        color: #333;
      }
      
      .welcome-text {
        font-size: 16px;
        line-height: 1.6;
        max-width: 600px;
        margin: 0 auto;
      }
    "))
  ),
  
  # 主页面布局
  div(class = "title-panel",
      h1("期权隐含波动率及历史波动率锥分析工具", class = "custom-title")
  ),
  
  sidebarLayout(
    # 侧边栏
    sidebarPanel(
      class = "sidebar-panel",
      width = 3,
      
      # 选择期权类型
      div(class = "select-input",
          selectInput("type", 
                      label = div(icon("chart-line"), "选择期权类型:"),
                      choices = c("看涨期权" = "call", "看跌期权" = "put"),
                      selected = "call")
      ),
      
      # 文件上传
      div(class = "file-input",
          fileInput("file", 
                    label = div(icon("file-upload"), "上传 Excel 文件"),
                    accept = c(".xlsx"),
                    buttonLabel = "浏览...",
                    placeholder = "选择.xlsx文件")
      ),
      
      # 处理按钮
      div(class = "action-button-container",
          actionButton("process", 
                       label = div(icon("refresh"), "处理数据并更新"),
                       class = "action-button")
      ),
      
      # 保存数据按钮
      div(class = "action-button-container",
          actionButton("save_data", 
                       label = div(icon("save"), "保存数据到本地"),
                       class = "save-button")
      ),
      
      # 进度条显示
      uiOutput("progress_ui"),
      
      # 分隔线
      div(class = "custom-hr"),
      
      # 信息提示
      div(style = "margin-top: 20px; font-size: 12px; color: #666;",
          icon("info-circle"),
          "上传Excel文件后点击处理按钮更新数据，检查无误后点击保存按钮"
      )
    ),
    
    # 主面板
    mainPanel(
      class = "main-panel",
      width = 9,
      
      # 添加欢迎信息
      uiOutput("welcome_ui"),
      
      # 只在有数据时显示标签页
      conditionalPanel(
        condition = "output.data_available",
        tabsetPanel(
          id = "main_tabs",
          type = "pills",
          
          # IV趋势图标签页
          tabPanel(
            title = div(icon("line-chart"), "IV 趋势图"),
            value = "iv_trend",
            div(class = "plot-container",
                plotlyOutput("iv_plot", height = "600px")
            )
          ),
          
          # 标的历史波动率标签页
          tabPanel(
            title = div(icon("bar-chart"), "标的历史波动率"),
            value = "iv_cone",
            div(class = "plot-container",
                plotlyOutput("iv_cone_plot", height = "600px")
            )
          ),
          
          # 历史波动率锥标签页
          tabPanel(
            title = div(icon("area-chart"), "历史波动率锥"),
            value = "hv_cone",
            div(class = "plot-container",
                plotlyOutput("hv_cone_plot", height = "600px")
            )
          ),
          
          # 数据表标签页
          tabPanel(
            title = div(icon("table"), "数据表"),
            value = "data_table",
            div(class = "data-table-container",
                h4(icon("database"), "期权隐含波动率数据表"),
                DTOutput("iv_data")
            )
          )
        )
      )
    )
  ),
  
  # 页脚
  div(style = "text-align: center; margin-top: 30px; padding: 15px; color: #666; border-top: 1px solid #eee;",
      "期权隐含波动率分析工具 © 2025 | ",
      tags$a(href = "#", "使用说明", style = "color: #667eea;"),
      " | ",
      tags$a(href = "#", "关于我们", style = "color: #667eea;")
  )
)

# ----------------------
# Server 部分
# ----------------------
server <- function(input, output, session) {
  
  # 进度条状态
  progress_status <- reactiveVal("")
  progress_value <- reactiveVal(0)
  total_steps <- 22
  
  # 数据可用性状态
  data_available <- reactiveVal(FALSE)
  
  # 进度条UI
  output$progress_ui <- renderUI({
    if (progress_status() != "") {
      div(class = "progress-container",
          div(class = "progress",
              div(class = "progress-bar", 
                  style = paste0("width: ", progress_value(), "%;"))
          ),
          div(class = "progress-text", 
              paste0(progress_status(), " (", round(progress_value()), "%)"))
      )
    }
  })
  
  # 更新进度条函数
  update_progress <- function(status, step) {
    progress_value((step / total_steps) * 100)
    progress_status(status)
  }
  
  # 欢迎信息UI
  output$welcome_ui <- renderUI({
    if (!data_available()) {
      div(class = "welcome-message",
          div(class = "welcome-icon", icon("upload")),
          div(class = "welcome-title", "欢迎使用期权波动率分析工具"),
          div(class = "welcome-text",
              p("请按照以下步骤开始分析："),
              p("1. 在左侧选择期权类型（看涨/看跌）"),
              p("2. 上传包含期权数据的Excel文件"),
              p("3. 点击『处理数据并更新』按钮"),
              p("4. 检查图表和数据表确认无误"),
              p("5. 点击『保存数据到本地』按钮")
          )
      )
    }
  })
  
  # 数据可用性输出
  output$data_available <- reactive({
    data_available()
  })
  outputOptions(output, "data_available", suspendWhenHidden = FALSE)
  
  # 修改：创建一个reactive value来存储当前处理的数据
  current_data <- reactiveVal(NULL)
  
  # 修改：加载历史数据的函数，现在会同时考虑文件数据和当前处理的数据
  get_historical_data <- reactive({
    req(input$type)
    
    # 如果当前有处理好的数据，优先使用
    if (!is.null(current_data())) {
      return(current_data())
    }
    
    # 否则从文件加载（但只在data_available为TRUE时）
    if (data_available()) {
      if (input$type == "call") {
        if (file.exists("call_data.RData")) {
          load("call_data.RData")
          return(list(
            result_list_latest = result_list_latest,
            atm_max_vol_df_latest = atm_max_vol_df_latest
          ))
        }
      } else if (input$type == "put") {
        if (file.exists("put_data.RData")) {
          load("put_data.RData")
          return(list(
            result_list_latest = result_list_latest,
            atm_max_vol_df_latest = atm_max_vol_df_latest
          ))
        }
      }
    }
    return(NULL)
  })
  
  # Black-Scholes 函数定义
  bs_call_price <- function(S, K, r, T, sigma) {
    d1 <- (log(S/K) + (r + 0.5*sigma^2)*T) / (sigma*sqrt(T))
    d2 <- d1 - sigma*sqrt(T)
    C <- S * pnorm(d1) - K * exp(-r*T) * pnorm(d2)
    return(C)
  }
  
  implied_vol_call <- function(C_market, S, K, r, T, sigma_lower=1e-6, sigma_upper=5) {
    if (C_market <= max(S - K * exp(-r*T), 0) || C_market > S) {
      return(NA)  
    }
    
    tryCatch({
      uniroot(function(sigma) bs_call_price(S, K, r, T, sigma) - C_market, 
              lower=sigma_lower, upper=sigma_upper)$root
    }, error=function(e) {
      message(sprintf("IV calculation failed for S=%.4f, K=%.4f, T=%.4f, C=%.4f, r=%.4f: %s", 
                      S, K, T, C_market, r, e$message))
      NA
    })
  }
  
  calculate_iv_call <- function(df) {
    df %>%
      mutate(
        Date = as.Date(Date),
        Expire_Date = as.Date(Expire_Date),
        T = as.numeric(difftime(Expire_Date, Date, units = "days")) / 365,
        IV = mapply(
          function(c_market, S, K, r, T_val) {
            if (is.na(c_market) || is.na(S) || is.na(K) || is.na(r) || is.na(T_val) || 
                T_val <= 0 || S <= 0 || K <= 0 || c_market <= 0) {
              return(NA)
            }
            implied_vol_call(c_market, S, K, r/100, T_val)
          },
          Close_Price, Target_Close_Price, Exercise_Price, rf, T
        )
      ) %>%
      select(-T)
  }
  
  bs_put_price <- function(S, K, r, T, sigma) {
    d1 <- (log(S/K) + (r + 0.5*sigma^2)*T) / (sigma*sqrt(T))
    d2 <- d1 - sigma*sqrt(T)
    P <- K * exp(-r*T) * pnorm(-d2) - S * pnorm(-d1)
    return(P)
  }
  
  implied_vol_put <- function(P_market, S, K, r, T, sigma_lower=1e-6, sigma_upper=5) {
    if (P_market <= max(K * exp(-r*T) - S, 0) || P_market > S) {
      return(NA)
    }
    
    tryCatch({
      uniroot(function(sigma) bs_put_price(S, K, r, T, sigma) - P_market, 
              lower=sigma_lower, upper=sigma_upper)$root
    }, error=function(e) {
      message(sprintf("IV calculation failed for S=%.4f, K=%.4f, T=%.4f, P=%.4f, r=%.4f: %s", 
                      S, K, T, P_market, r, e$message))
      NA
    })
  }
  
  calculate_iv_put <- function(df) {
    df %>%
      mutate(
        Date = as.Date(Date),
        Expire_Date = as.Date(Expire_Date),
        T = as.numeric(difftime(Expire_Date, Date, units = "days")) / 365,
        IV = mapply(
          function(p_market, S, K, r, T_val) {
            if (is.na(p_market) || is.na(S) || is.na(K) || is.na(r) || is.na(T_val) || 
                T_val <= 0 || S <= 0 || K <= 0 || p_market <= 0) {
              return(NA)
            }
            implied_vol_put(p_market, S, K, r/100, T_val)
          },
          Close_Price, Target_Close_Price, Exercise_Price, rf, T
        )
      ) %>%
      select(-T)
  }
  
  # 1️⃣ 数据读取和清洗
  data_processed <- eventReactive(input$process, {
    req(input$file, input$type)
    
    # 设置数据可用状态
    data_available(TRUE)
    
    # 重置进度条
    update_progress("开始处理数据...", 1)
    
    # ----------------------
    # 读取并清洗函数
    # ----------------------
    clean <- function(df) {
      df[1, 1] <- df[2, 1]
      df <- df[-2, ]
      colnames(df) <- as.character(unlist(df[1, ]))
      df <- df[-1, ]
      df[[1]] <- as.Date(as.numeric(df[[1]]), origin = "1899-12-30")
      if(ncol(df) > 1) df[,-1] <- lapply(df[,-1], function(x) suppressWarnings(as.numeric(x)))
      df
    }
    
    update_progress("读取Excel文件...", 2)
    
    file_path <- input$file$datapath
    
    # 使用进度条封装读取过程
    withProgress(message = '读取数据表...', value = 0, {
      close_price <- suppressMessages(read_excel(file_path, sheet = "收盘价", col_names = FALSE, skip = 2))
      incProgress(1/6)
      exercise_price <- suppressMessages(read_excel(file_path, sheet = "行权价格", col_names = FALSE, skip = 2))
      incProgress(1/6)
      target_close_price <- suppressMessages(read_excel(file_path, sheet = "标的收盘价", col_names = FALSE, skip = 2))
      incProgress(1/6)
      expire_date <- suppressMessages(read_excel(file_path, sheet = "到期日", col_names = FALSE, skip = 2))
      incProgress(1/6)
      volume <- suppressMessages(read_excel(file_path, sheet = "成交量", col_names = FALSE, skip = 2))
      incProgress(1/6)
      rf <- suppressMessages(read_excel(file_path, sheet = "无风险利率", col_names = FALSE, skip = 4))
      incProgress(1/6)
    })
    
    update_progress("数据清洗中...", 3)
    
    ID <- data.frame(ID = t(close_price[1, -1]), stringsAsFactors = FALSE)
    
    rf$时间 <- as.Date(rf[[1]])
    colnames(rf) <- c("时间", "无风险利率")
    rf[,-1] <- lapply(rf[,-1], as.numeric)
    
    close_price <- clean(close_price)
    exercise_price <- clean(exercise_price)
    target_close_price <- clean(target_close_price)
    expire_date <- clean(expire_date)
    volume <- clean(volume)
    
    expire_date[-1] <- lapply(expire_date[-1], function(col) as.Date(as.numeric(col), origin = "1899-12-30"))
    
    if(!all(nrow(volume) == nrow(close_price), nrow(volume) == nrow(target_close_price),
            nrow(volume) == nrow(expire_date), nrow(volume) == nrow(exercise_price))) {
      stop("所有数据的行数必须相同")
    }
    
    if(!all(ncol(volume)-1 == nrow(ID))) stop("所有数据的ID必须相同")
    
    # ----------------------
    # ID 添加到期日和行权价
    # ----------------------
    update_progress("处理期权ID信息...", 4)
    
    ids <- as.character(ID[[1]])
    ID$Expire_Date <- as.Date(NA)
    ID$Exercise_Price <- NA_real_
    for(i in seq_along(ids)) {
      id <- ids[i]
      if(id %in% colnames(expire_date)) {
        expiry <- na.omit(expire_date[[id]])
        if(length(expiry) > 0) {
          ID$Expire_Date[i] <- expiry[1]
        }
      }
      if(id %in% colnames(exercise_price)) {
        exercise_price_vals <- na.omit(exercise_price[[id]])
        if(length(exercise_price_vals) > 0) {
          ID$Exercise_Price[i] <- exercise_price_vals[1]
        }
      }
    }
    
    for(id in ids) {
      if(id %in% colnames(close_price) && id %in% colnames(volume)) {
        expiry_date <- ID$Expire_Date[ID[[1]] == id]
        
        if(!is.na(expiry_date)) {
          expire_after_indices <- which(close_price$时间 > expiry_date)
          
          if(length(expire_after_indices) > 0) {
            close_price[expire_after_indices, id] <- 0
            volume[expire_after_indices, id] <- 0
          }
        }
      }
    }
    
    # ----------------------
    # 筛选缺失日期
    # ----------------------
    update_progress("筛选有效日期数据...", 5)
    
    # 获取历史数据以确定最新日期
    hist_data <- get_historical_data()
    latest_date <- if(!is.null(hist_data$result_list_latest)) {
      max(as.Date(names(hist_data$result_list_latest)))
    } else {
      max(close_price$时间, na.rm = TRUE)
    }
    
    today <- Sys.Date()
    if (latest_date + 1 > today - 1) {
      diff_dates <- close_price$时间[close_price$时间 != today]
    } else {
      diff_dates <- seq(from = latest_date + 1, to = today - 1, by = "day")
      diff_dates <- diff_dates[diff_dates %in% volume$时间]
    }
    
    # ----------------------
    # 构建每日数据 update_list
    # ----------------------
    update_progress("构建每日数据...", 6)
    
    update_list <- list()
    total_rows <- nrow(volume)
    
    withProgress(message = '构建每日数据...', value = 0, {
      for(i in 1:total_rows) {
        current_date <- volume$时间[i]
        if(current_date %in% diff_dates) {
          volume_values <- as.numeric(volume[i, -1])
          close_price_values <- as.numeric(close_price[i, -1])
          target_close_price_values <- as.numeric(target_close_price[i, -1])
          
          expire_date_values <- ID$Expire_Date[match(ids, ID[[1]])]
          exercise_price_values <- ID$Exercise_Price[match(ids, ID[[1]])]
          
          date_data <- data.frame(
            ID = ids,
            Volume = volume_values,
            Close_Price = close_price_values,
            Target_Close_Price = target_close_price_values,
            Expire_Date = expire_date_values,
            Exercise_Price = exercise_price_values,
            stringsAsFactors = FALSE
          )
          
          date_data <- subset(
            date_data,
            !( (is.na(Volume) | Volume == 0) & (is.na(Close_Price) | Close_Price == 0) )
          )
          
          if(nrow(date_data) > 0 && !all(is.na(date_data[, -1]))) {
            update_list[[as.character(current_date)]] <- date_data
          }
        }
        incProgress(1/total_rows)
      }
    })
    
    # ----------------------
    # 添加 rf
    # ----------------------
    update_progress("添加无风险利率数据...", 7)
    
    rf$时间 <- as.Date(rf$时间)
    for(date_str in names(update_list)) {
      current_date <- as.Date(date_str)
      rf_row <- rf[rf$时间 == current_date, ]
      update_list[[date_str]]$rf <- if(nrow(rf_row) > 0) rf_row$无风险利率[1] else NA
    }
    
    # ----------------------
    # 计算 Level
    # ----------------------
    update_progress("计算期权Level...", 8)
    
    for(date_str in names(update_list)) {
      date_data <- update_list[[date_str]]
      date_data <- date_data %>% mutate(diff = abs(Target_Close_Price - Exercise_Price))
      min_diff_idx <- which.min(date_data$diff)
      date_data$level <- 0
      at_the_money_strike <- date_data$Exercise_Price[min_diff_idx]
      greater_than_atm <- date_data %>% filter(Exercise_Price > at_the_money_strike) %>% arrange(Exercise_Price) %>% mutate(level = dense_rank(Exercise_Price))
      less_than_atm <- date_data %>% filter(Exercise_Price < at_the_money_strike) %>% arrange(desc(Exercise_Price)) %>% mutate(level = -dense_rank(desc(Exercise_Price)))
      at_the_money <- date_data %>% filter(Exercise_Price == at_the_money_strike)
      date_data <- bind_rows(greater_than_atm, less_than_atm, at_the_money)
      update_list[[date_str]] <- date_data
    }
    
    # ----------------------
    # 筛选掉临近到期日的合约
    # ----------------------
    update_progress("筛选临近到期合约...", 9)
    
    for (i in seq_along(update_list)) {
      current_date <- as.Date(names(update_list)[i])
      df <- update_list[[i]]
      df$Expire_Date <- as.Date(df$Expire_Date)
      df <- df[df$Expire_Date - current_date > 1, ]
      update_list[[i]] <- df
    }
    
    # ----------------------
    # 提取平值期权
    # ----------------------
    update_progress("提取平值期权...", 10)
    
    At_the_money <- lapply(update_list, function(date_data) {
      date_data <- date_data %>%
        mutate(diff = abs(Target_Close_Price - Exercise_Price))
      min_diff <- min(date_data$diff, na.rm = TRUE)
      date_data %>% filter(diff == min_diff)
    })
    
    # ----------------------
    # 按每日最大成交量筛选主力合约
    # ----------------------
    update_progress("筛选主力合约...", 11)
    
    atm_max_vol <- lapply(At_the_money, function(date_data) {
      if(nrow(date_data) == 0) return(date_data[0, ])
      max_vol <- max(date_data$Volume, na.rm = TRUE)
      date_data %>% filter(Volume == max_vol)
    })
    
    atm_max_vol_df <- bind_rows(atm_max_vol, .id = "Date") %>%
      select(-diff)
    
    # ----------------------
    # 计算隐含波动率
    # ----------------------
    update_progress("计算隐含波动率...", 12)
    
    if (input$type == "call"){
      atm_max_vol_df <- calculate_iv_call(atm_max_vol_df)
    } else if (input$type == "put") {
      atm_max_vol_df <- calculate_iv_put(atm_max_vol_df)
    }
    
    # ----------------------
    # 合并历史数据 - update_list
    # ----------------------
    update_progress("合并历史数据...", 13)
    
    hist_data <- get_historical_data()
    if (!is.null(hist_data) && !is.null(hist_data$result_list_latest)) {
      dates_A <- names(hist_data$result_list_latest)
      dates_B <- names(update_list)
      all_dates <- sort(unique(c(dates_A, dates_B)), decreasing = TRUE) 
      result_list <- list()
      
      for (d in all_dates) {
        if (d %in% dates_A & d %in% dates_B) {
          df_A <- hist_data$result_list_latest[[d]]
          df_B <- update_list[[d]]
          merged_df <- bind_rows(df_A, df_B) %>% distinct()
          result_list[[d]] <- merged_df
        } else if (d %in% dates_A) {
          result_list[[d]] <- hist_data$result_list_latest[[d]]
        } else {
          result_list[[d]] <- update_list[[d]]
        }
      }
      
      for (d in names(result_list)) {
        result_list[[d]] <- result_list[[d]] %>% filter(if_any(everything(), ~ !is.na(.)))
      }
      
      update_list <- result_list
    }
    
    # ----------------------
    # 合并历史数据
    # ----------------------
    update_progress("合并主力合约数据...", 14)
    
    if (!is.null(hist_data) && !is.null(hist_data$atm_max_vol_df_latest)) {
      atm_max_vol_df <- bind_rows(atm_max_vol_df, hist_data$atm_max_vol_df_latest) %>%
        distinct() %>%
        group_by(Date) %>%
        slice_max(Volume, n = 1, with_ties = FALSE) %>% 
        ungroup() %>%
        arrange(desc(Date)) %>%
        filter(if_any(everything(), ~ !is.na(.)))
    }
    
    update_progress("数据处理完成！", 15)
    
    return(list(
      update_list = update_list, 
      atm_max_vol_df = atm_max_vol_df,
      dates = sort(as.Date(names(update_list)))
    ))
  })
  
  # 2️⃣ 处理平值期权 IV 计算
  atm_iv_data <- eventReactive(input$process, {
    req(data_processed())
    
    update_progress("开始计算隐含波动率...", 16)
    
    data <- data_processed()
    update_list <- data$update_list
    atm_max_vol_df <- data$atm_max_vol_df
    type <- input$type
    
    update_progress("IV计算完成！", 17)
    
    return(list(
      At_the_money = update_list,
      atm_max_vol_df = atm_max_vol_df
    ))
  })
  
  # 3️⃣ 处理数据按钮 - 只更新current_data，不保存文件
  observeEvent(input$process, {
    req(atm_iv_data())
    
    update_progress("准备数据展示...", 18)
    
    data <- atm_iv_data()
    result_list_latest <- data$At_the_money
    atm_max_vol_df_latest <- data$atm_max_vol_df
    
    # 只更新current_data，不保存到文件
    current_data(list(
      result_list_latest = result_list_latest,
      atm_max_vol_df_latest = atm_max_vol_df_latest
    ))
    
    update_progress("数据处理完成！", 19)
    
    # 显示成功通知，提示用户检查后保存
    showNotification("数据处理完成！请检查图表和数据表，确认无误后点击『保存数据到本地』按钮。", 
                     type = "message", duration = 10)
  })
  
  # 4️⃣ 保存数据按钮的逻辑
  observeEvent(input$save_data, {
    req(current_data())
    
    # 显示保存进度
    save_notification <- showNotification("正在保存数据到本地文件...", 
                                          type = "message", 
                                          duration = NULL)
    
    data <- current_data()
    result_list_latest <- data$result_list_latest
    atm_max_vol_df_latest <- data$atm_max_vol_df_latest
    
    # 保存到RData和Excel文件
    if (input$type == "call"){
      write_xlsx(atm_max_vol_df_latest, path = "IV_call_results.xlsx")
      save(result_list_latest, atm_max_vol_df_latest, file = "call_data.RData")
      showNotification("Call数据已成功保存到本地！", type = "message", duration = 5)
    } else if (input$type == "put") {
      write_xlsx(atm_max_vol_df_latest, path = "IV_put_results.xlsx")
      save(result_list_latest, atm_max_vol_df_latest, file = "put_data.RData")
      showNotification("Put数据已成功保存到本地！", type = "message", duration = 5)
    }
    
    # 移除保存通知
    removeNotification(save_notification)
  })
  
  # 修改：创建一个统一的reactive来获取当前显示的atm_max_vol_df
  display_atm_max_vol_df <- reactive({
    # 只在数据可用时返回数据
    if (!data_available()) {
      return(data.frame())
    }
    
    # 优先使用当前处理的数据
    current <- current_data()
    if (!is.null(current) && !is.null(current$atm_max_vol_df_latest)) {
      return(current$atm_max_vol_df_latest)
    }
    
    # 否则使用历史数据
    hist_data <- get_historical_data()
    if (!is.null(hist_data) && !is.null(hist_data$atm_max_vol_df_latest)) {
      return(hist_data$atm_max_vol_df_latest)
    }
    
    # 如果没有数据，返回空数据框
    return(data.frame())
  })
  
  # 5️⃣ IV 趋势图
  output$iv_plot <- renderPlotly({
    req(display_atm_max_vol_df(), nrow(display_atm_max_vol_df()) > 0)
    
    update_progress("生成IV趋势图...", 20)
    
    atm_max_vol_df <- display_atm_max_vol_df()
    
    # 确保日期格式正确
    atm_max_vol_df$Date <- as.Date(atm_max_vol_df$Date)
    
    # 创建交互式图表
    fig <- plot_ly(
      data = atm_max_vol_df,
      x = ~Date,
      y = ~IV,
      type = 'scatter',
      mode = 'lines+markers',
      line = list(color = 'steelblue', width = 2),
      marker = list(size = 5, color = 'steelblue', opacity = 0.8),
      hovertemplate = paste(
        "<b>日期:</b> %{x}<br>",
        "<b>隐含波动率:</b> %{y:.2%}<extra></extra>"
      )
    )
    
    # 添加布局 
    fig <- fig %>%
      layout(
        title = list(
          text = "平值期权隐含波动率趋势",
          x = 0.5,
          y = 0.95,
          font = list(size = 20, family = "Microsoft YaHei", color = "#333")
        ),
        xaxis = list(
          title = "日期",
          type = "date",
          tickformat = "%Y-%m-%d",
          rangeslider = list(visible = TRUE),
          rangeselector = list(
            buttons = list(
              list(count = 7, label = "1周", step = "day", stepmode = "backward"),
              list(count = 1, label = "1月", step = "month", stepmode = "backward"),
              list(count = 3, label = "3月", step = "month", stepmode = "backward"),
              list(count = 6, label = "半年", step = "month", stepmode = "backward"),
              list(step = "all", label = "全部")
            )
          )
        ),
        yaxis = list(
          title = "隐含波动率",
          tickformat = ".1%",
          rangemode = "tozero"
        ),
        hovermode = "x unified",
        plot_bgcolor = "#fafafa",
        paper_bgcolor = "#ffffff",
        margin = list(t = 100, b = 80, l = 80, r = 40) 
      )
    
    fig
  })
  
  # 6️⃣ 标的历史波动率
  output$iv_cone_plot <- renderPlotly({
    req(display_atm_max_vol_df(), nrow(display_atm_max_vol_df()) > 0)
    
    update_progress("生成标的历史波动率...", 21)
    
    # 使用display_atm_max_vol_df()来获取数据
    atm_max_vol_df <- display_atm_max_vol_df()
    
    # 使用已有的数据计算历史波动率
    price_df <- atm_max_vol_df %>%
      arrange(Date) %>%
      mutate(Return = log(Target_Close_Price / lag(Target_Close_Price))) %>%
      filter(!is.na(Return))
    
    window_list <- c(10, 20, 60, 120)
    
    # 计算不同窗口期的历史波动率
    hv_data <- data.frame()
    for (w in window_list) {
      hv_tmp <- price_df %>%
        mutate(
          HV = rollapply(Return, w, sd, fill = NA, align = "right") * sqrt(252),
          Window = paste0(w, "天")
        )
      hv_data <- bind_rows(hv_data, hv_tmp)
    }
    
    hv_data <- hv_data %>% filter(!is.na(HV))
    
    # 绘制历史波动率时间序列图
    fig <- plot_ly() 
    
    # 为每个窗口期添加一条线
    colors <- c("#1f77b4", "#ff7f0e", "#2ca02c", "#d62728")
    
    for (i in 1:length(window_list)) {
      w <- window_list[i]
      window_data <- hv_data %>% filter(Window == paste0(w, "天"))
      
      fig <- fig %>%
        add_trace(
          data = window_data,
          x = ~Date,
          y = ~HV,
          type = 'scatter',
          mode = 'lines',
          name = paste0(w, "天历史波动率"),
          line = list(color = colors[i], width = 2),
          hovertemplate = paste(
            "<b>日期:</b> %{x}<br>",
            "<b>", w, "天HV:</b> %{y:.2%}<extra></extra>"
          )
        )
    }
    
    # 添加当前日期标记
    latest_date <- max(hv_data$Date, na.rm = TRUE)
    latest_hv <- hv_data %>% 
      filter(Date == latest_date) %>%
      arrange(Window)
    
    for (i in 1:nrow(latest_hv)) {
      fig <- fig %>%
        add_trace(
          x = ~c(latest_hv$Date[i]),
          y = ~c(latest_hv$HV[i]),
          type = 'scatter',
          mode = 'markers',
          marker = list(
            size = 8, 
            color = colors[i],
            symbol = "diamond"
          ),
          name = paste0("当前", latest_hv$Window[i]),
          showlegend = FALSE,
          hovertemplate = paste(
            "<b>日期:</b>", latest_date, "<br>",
            "<b>", gsub("天", "", latest_hv$Window[i]), "天HV:</b> %{y:.2%}<extra></extra>"
          )
        )
    }
    
    fig <- fig %>%
      layout(
        title = list(
          text = "历史波动率时间序列",
          x = 0.5,
          font = list(size = 20, family = "Microsoft YaHei", color = "#333")
        ),
        xaxis = list(
          title = "日期",
          type = "date",
          tickformat = "%Y-%m-%d",
          rangeslider = list(visible = TRUE),
          rangeselector = list(
            buttons = list(
              list(count = 1, label = "1月", step = "month", stepmode = "backward"),
              list(count = 3, label = "3月", step = "month", stepmode = "backward"),
              list(count = 6, label = "6月", step = "month", stepmode = "backward"),
              list(count = 1, label = "1年", step = "year", stepmode = "backward"),
              list(step = "all", label = "全部")
            )
          )
        ),
        yaxis = list(
          title = "年化历史波动率",
          tickformat = ".1%",
          rangemode = "tozero",
          domain = c(0.0, 1)
        ),
        hovermode = "x unified",
        legend = list(
          orientation = "h", 
          x = 0.5, 
          y = 1.02, 
          xanchor = "center",
          yanchor = "bottom",
          bgcolor = "rgba(255,255,255,0.9)",
          bordercolor = "rgba(0,0,0,0.1)",
          borderwidth = 1,
          font = list(size = 12, family = "Microsoft YaHei")
        ),
        plot_bgcolor = "#fafafa",
        paper_bgcolor = "#ffffff",
        margin = list(t = 120, b = 120)
      )
    
    fig
  })
  
  # 7️⃣ 历史波动率锥
  output$hv_cone_plot <- renderPlotly({
    req(display_atm_max_vol_df(), nrow(display_atm_max_vol_df()) > 0)
    
    update_progress("生成历史波动率锥...", 22)
    
    # 处理完成后隐藏进度条
    Sys.sleep(0.5)
    update_progress("", 0)
    
    atm_max_vol_df <- display_atm_max_vol_df()
    
    # Step 1: 计算每日对数收益率
    price_df <- atm_max_vol_df %>%
      arrange(Date) %>%
      mutate(Return = log(Target_Close_Price / lag(Target_Close_Price))) %>%
      filter(!is.na(Return))
    
    # Step 2: 定义滚动窗口
    window_list <- c(10, 20, 60, 120)
    
    # Step 3: 计算滚动年化历史波动率
    hv_data <- data.frame()
    for (w in window_list) {
      hv_tmp <- price_df %>%
        mutate(
          HV = rollapply(Return, w, sd, fill = NA, align = "right") * sqrt(252),
          Window = w
        )
      hv_data <- bind_rows(hv_data, hv_tmp)
    }
    
    hv_data <- hv_data %>% filter(!is.na(HV))
    
    # Step 4: 静态分位数（用所有数据计算一次）
    hv_cone_static <- hv_data %>%
      group_by(Window) %>%
      summarise(
        p5 = quantile(HV, 0.05, na.rm = TRUE),
        p25 = quantile(HV, 0.25, na.rm = TRUE),
        median = quantile(HV, 0.5, na.rm = TRUE),
        p75 = quantile(HV, 0.75, na.rm = TRUE),
        p95 = quantile(HV, 0.95, na.rm = TRUE)
      )
    
    # Step 5: 动画红点数据
    hv_points <- hv_data %>%
      mutate(frame = as.character(Date)) %>%
      select(Window, HV, frame)
    
    # Step 6: 绘图优化版
    p <- plot_ly() %>%
      # 灰色区间（5%-95%）
      add_trace(
        data = hv_cone_static,
        x = ~Window, y = ~p95,
        type = "scatter", mode = "lines",
        line = list(color = "lightgray"),
        name = "95分位", showlegend = FALSE
      ) %>%
      add_trace(
        data = hv_cone_static,
        x = ~Window, y = ~p5,
        type = "scatter", mode = "lines",
        fill = "tonexty", fillcolor = "rgba(211,211,211,0.3)",
        line = list(color = "lightgray"),
        name = "5-95分位区间", showlegend = TRUE
      ) %>%
      # 蓝色区间（25%-75%）
      add_trace(
        data = hv_cone_static,
        x = ~Window, y = ~p75,
        type = "scatter", mode = "lines",
        line = list(color = "skyblue"), name = "75分位", showlegend = FALSE
      ) %>%
      add_trace(
        data = hv_cone_static,
        x = ~Window, y = ~p25,
        type = "scatter", mode = "lines",
        fill = "tonexty", fillcolor = "rgba(135,206,235,0.4)",
        line = list(color = "skyblue"),
        name = "25-75分位区间", showlegend = TRUE
      ) %>%
      # 中位数线
      add_trace(
        data = hv_cone_static,
        x = ~Window, y = ~median,
        type = "scatter", mode = "lines+markers",
        line = list(color = "steelblue", width = 3),
        marker = list(size = 5, color = "steelblue"),
        name = "中位数"
      ) %>%
      # 红点动画
      add_trace(
        data = hv_points,
        x = ~Window, y = ~HV,
        frame = ~frame,
        type = "scatter", mode = "markers+text",
        marker = list(size = 8, color = "red"),
        text = ~paste0(Window, "天HV: ", round(HV * 100, 2), "%"),
        textposition = "top center",
        hovertemplate = paste0(
          "窗口: %{x}天<br>",
          "日期: %{frame}<br>",
          "HV: %{y:.2%}<br><extra></extra>"
        ),
        name = "目标日期HV"
      ) %>%
      layout(
        title = list(
          text = "历史波动率锥",
          x = 0.5,
          y = 0.95,
          font = list(size = 16, family = "Microsoft YaHei", color = "#333")
        ),
        xaxis = list(
          title = "滚动窗口长度（天）",
          tickvals = window_list,
          ticktext = paste0(window_list, "天")
        ),
        yaxis = list(
          title = "年化历史波动率 (HV)",
          tickformat = ".1%",
          range = c(0, 1),
          domain = c(0.0, 1)
        ),
        hovermode = "x unified",
        legend = list(
          orientation = "h",
          x = 0.5,
          xanchor = "center",
          y = -0.25 
        ),
        plot_bgcolor = "#fafafa",
        paper_bgcolor = "#ffffff",
        margin = list(t = 100, b = 150, l = 80, r = 40)
      ) %>%
      animation_opts(
        frame = 100,
        transition = 0,
        easing = "linear",
        redraw = FALSE
      ) %>%
      animation_slider(
        currentvalue = list(prefix = "日期: ", font = list(color = "black")),
        pad = list(t = 100, b = 10)
      )
    
    p
  })
  
  # 8️⃣ 数据表显示 - 修改为使用display_atm_max_vol_df
  output$iv_data <- renderDT({
    req(display_atm_max_vol_df(), nrow(display_atm_max_vol_df()) > 0)
    
    atm_max_vol_df <- display_atm_max_vol_df()
    
    # 格式化数据，使显示更友好
    formatted_data <- atm_max_vol_df %>%
      mutate(
        Date = format(Date, "%Y-%m-%d"),
        Expire_Date = format(Expire_Date, "%Y-%m-%d"),
        IV = round(IV * 100, 2),
        Close_Price = round(Close_Price, 4),
        Target_Close_Price = round(Target_Close_Price, 4),
        Exercise_Price = round(Exercise_Price, 4),
        rf = round(rf, 4)
      ) %>%
      rename(
        "交易日期" = Date,
        "期权代码" = ID,
        "成交量" = Volume,
        "期权收盘价" = Close_Price,
        "标的收盘价" = Target_Close_Price,
        "到期日" = Expire_Date,
        "行权价" = Exercise_Price,
        "无风险利率(%)" = rf,
        "隐含波动率(%)" = IV
      )
    
    # 创建数据表
    datatable(
      formatted_data,
      options = list(
        pageLength = 20,
        lengthMenu = c(10, 20, 50, 100),
        scrollX = TRUE,
        autoWidth = TRUE,
        dom = 'Bfrtip',
        buttons = c('copy', 'csv', 'excel', 'pdf', 'print')
      ),
      extensions = 'Buttons',
      rownames = FALSE,
      filter = 'top',
      class = 'cell-border stripe hover'
    ) %>%
      formatStyle(
        '隐含波动率(%)',
        background = styleColorBar(range(formatted_data$`隐含波动率(%)`, na.rm = TRUE), 'lightblue'),
        backgroundSize = '98% 88%',
        backgroundRepeat = 'no-repeat',
        backgroundPosition = 'center'
      )
  })
  
}

# ----------------------  
# 启动 Shiny App
# ----------------------
shinyApp(ui, server)