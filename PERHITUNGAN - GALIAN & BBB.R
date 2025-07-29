#LIBRARY
library(dplyr)
library(tidyr)
library(readxl)
library(writexl)
library(openxlsx)
library(rlang)

#INPUT FILES

input_files <- "Input/InputValue.v.1.xlsx"

galian_pilecap <- read_excel(input_files, sheet = "GALIAN - PILE CAP")
galian_sloof <- read_excel(input_files, sheet = "GALIAN - SLOOF")
massa_besi <- read_excel(input_files, sheet = "Tabel Berat Besi")
bbb_pilecap <- read_excel(input_files, sheet = "BBB - A. PILE CAP")
bbb_sloof <- read_excel(input_files, sheet = "BBB - B. SLOOF")
bbb_kolom_lantai <- read_excel(input_files, sheet = "BBB - C. KOLOM LANTAI")
bbb_pitlift <- read_excel(input_files, sheet = "BBB - D. PITLIFT")
bbb_kolom_lift <- read_excel(input_files, sheet = "BBB - D. KOLOM LIFT")
bbb_balok_lift <- read_excel(input_files, sheet = "BBB - D. BALOK LIFT")
bbb_balok_lantai <- read_excel(input_files, sheet = "BBB - E. BALOK LANTAI")
bbb_borepile <- read_excel(input_files, sheet = "BBB - F. BOREPILE")
bbb_tangga <- read_excel(input_files, sheet = "BBB - G. TANGGA")
bbb_plat1 <- read_excel(input_files, sheet = "BBB - H. PLAT1")
bbb_plat2 <- read_excel(input_files, sheet = "BBB - H. PLAT2")

NA_to_0 <- function(df){
  df[is.na(df)] <- 0
  return(df)
}

galian_pilecap <- NA_to_0(galian_pilecap)
galian_sloof<- NA_to_0(galian_sloof)
massa_besi<- NA_to_0(massa_besi)
bbb_pilecap<- NA_to_0(bbb_pilecap)
bbb_sloof<- NA_to_0(bbb_sloof)
bbb_kolom_lantai<- NA_to_0(bbb_kolom_lantai)
bbb_pitlift<- NA_to_0(bbb_pitlift)
bbb_kolom_lift<- NA_to_0(bbb_kolom_lift)
bbb_balok_lift<- NA_to_0(bbb_balok_lift)
bbb_balok_lantai<- NA_to_0(bbb_balok_lantai)
bbb_borepile<- NA_to_0(bbb_borepile)
bbb_tangga<- NA_to_0(bbb_tangga)
bbb_plat1<- NA_to_0(bbb_plat1)
bbb_plat2<- NA_to_0(bbb_plat2)

#DIAMETER TABLE
diameter <- c(
              bbb_pilecap$`Diameter Tulangan Atas (mm)`, bbb_pilecap$`Diameter Bawah (mm)`, bbb_pilecap$`Diameter Sabuk (mm)`,
              bbb_sloof$`Diameter Tulangan Utama (mm)`, bbb_sloof$`Diameter Begel (mm)`, bbb_sloof$`Diameter Sabuk (mm)`,
              bbb_kolom_lantai$`Diameter Begel (mm)`,
              bbb_pitlift$`Diameter Besi X (mm)`, bbb_pitlift$`Diameter Besi Y(mm)`,
              bbb_kolom_lift$`Diameter Tulangan Utama (mm)`, bbb_kolom_lift$`Diameter Begel (mm)`,
              bbb_balok_lift$`Diameter Tulangan Utama (mm)`, bbb_balok_lift$`Diameter Begel (mm)`, bbb_balok_lift$`Diameter Sabuk (mm)`,
              bbb_balok_lantai$`Diameter Tulangan Utama (mm)`, bbb_balok_lantai$`Diameter Begel (mm)`, bbb_balok_lantai$`Diameter Sabuk (mm)`,
              bbb_borepile$`Diameter Tulangan Utama (mm)`, bbb_borepile$`Diameter Begel (mm)`,
              bbb_tangga$`Diameter Tulangan Utama (mm)`, bbb_tangga$`Diameter Tulangan Bagi (mm)`,
              bbb_plat2$`Diameter Besi (mm)`
              )

diameter_unique <- sort(na.omit(unique(diameter)), decreasing = TRUE)
diameter_unique <- diameter_unique[diameter_unique!=0]

#USEFUL FUNCTION

#IFERROR
iferror <- function(expr, fallback) {
  result <- tryCatch(expr, error = function(e) fallback)
  # Optional: handle Inf, NaN, or NA
  if (is.infinite(result) || is.nan(result) || is.na(result)) {
    fallback
  } else {
    result
  }
}

#DIAMETER TABLE
d_table <- function(df){
  for (i in diameter_unique){
    df[[paste0("D",i," - btg")]] <- 0
    df[[paste0("D",i," - kg")]] <- 0
  }
  df$`Total - btg` <- 0
  df$`Total - kg` <- 0
  return(df)
}

#TOTAL ROW
add_total_row <- function(source,df,isfloor){
  if (isfloor == 1){
    raw_col <- ncol(source) - 1
  }else{
    raw_col <- ncol(source)
  }
  col_df <- colnames(df)
  n_col <- length(col_df)
  total_row <- c("TOTAL",rep(NA,raw_col-1))
  for (i in (raw_col+1):n_col){
    temp_col <- col_df[i]
    total_row <- c(total_row, sum(df[[temp_col]]))
  }
  df <- rbind(df,total_row)
  # for (i in (raw_col+1):n_col){
  #   temp_col <- col_df[i]
  #   df[[temp_col]] <- as.numeric(df[[temp_col]])
  # }
  for (i in 2:n_col){
    temp_col <- col_df[i]
    if(temp_col != "Notes"){
      df[[temp_col]] <- as.numeric(df[[temp_col]])
    }
  }
  return(df)
}



#CALCULATION FUNCTION/PROCESS

#GALIAN - PILE CAP
galian_pilecap_step_1 <- galian_pilecap %>%
  mutate(`Galian (m3)` = (`Tinggi Pilecap (m)`+`Elevasi muka atas dari tanah asli (m)`)*(`Dimensi Pilecap P (m)`+0.4)*(`Dimensi Pilecap L (m)` + 0.4) * `Jumlah Pilecap (bh)`,
         `Pasir (m3)` = 0.1*(`Dimensi Pilecap P (m)` + 0.2)*(`Dimensi Pilecap L (m)` + 0.2)*`Jumlah Pilecap (bh)`,
         `Lantai Kerja (m3)` = 0.05*(`Dimensi Pilecap P (m)` + 0.2)*(`Dimensi Pilecap L (m)` + 0.2)*`Jumlah Pilecap (bh)`)

galian_pilecap_step_2 <- add_total_row(galian_pilecap,galian_pilecap_step_1,0)

#GALIAN - SLOOF
galian_sloof_step_1 <- galian_sloof %>%
  mutate(`Galian (m3)` = `Panjang Sloof (m)`*(`Dimensi Sloof H (m)` + `Elevasi Muka dari atas Tanah Asli (m)`)*(`Dimensi Sloof B (m)` + 0.3),
         `Pasir (m3)` = 0.1*(`Dimensi Sloof B (m)` + 0.2)*`Panjang Sloof (m)`,
         `Lantai Kerja (m3)` = 0.05*(`Dimensi Sloof B (m)` + 0.2)*`Panjang Sloof (m)`)

galian_sloof_step_2 <- add_total_row(galian_sloof, galian_sloof_step_1,0)

#BBB - PILE CAP
bbb_pilecap_step_1 <- bbb_pilecap %>%
  mutate(`Volume Beton (m3)` = `Jumlah Pilecap (bh)`*`Tinggi Pilecap (m)`*`Dimensi Pilecap P (m)`*`Dimensi Pilecap L (m)`,
         `Begesting (m2)` = (`Dimensi Pilecap L (m)`*2 + `Dimensi Pilecap P (m)`*2)*`Tinggi Pilecap (m)`*`Jumlah Pilecap (bh)`)

bbb_pilecap_step_2 <- d_table(bbb_pilecap_step_1)

for (j in 1:nrow(bbb_pilecap_step_2)){
  
  for (k in diameter_unique){
    
    if(bbb_pilecap_step_2$`Diameter Tulangan Atas (mm)`[j] == k){
      temp <- bbb_pilecap_step_2
      col_name <- paste0("D",k," - btg")
      prev_value <- temp[[col_name]][j]
      new_value <- ((temp$`Dimensi Pilecap L (m)`[j]-0.1+(temp$`Tinggi Pilecap (m)`[j]-0.1)*2)*(temp$`Dimensi Pilecap P (m)`[j]/(temp$`Jarak Tulangan Atas (mm)`[j]/1000)+1)+(temp$`Dimensi Pilecap P (m)`[j]-0.1+(temp$`Tinggi Pilecap (m)`[j]-0.1)*2)*(temp$`Dimensi Pilecap L (m)`[j]/(temp$`Jarak Tulangan Atas (mm)`[j]/1000)+1))*1.05*temp$`Jumlah Pilecap (bh)`[j]/12
      temp[[col_name]][j] <- prev_value+new_value
      bbb_pilecap_step_2 <- temp
    }
    
    if(bbb_pilecap_step_2$`Diameter Bawah (mm)`[j] == k){
      temp <- bbb_pilecap_step_2
      col_name <- paste0("D",k," - btg")
      prev_value <- temp[[col_name]][j]
      new_value <- ((temp$`Dimensi Pilecap L (m)`[j]-0.1+(temp$`Tinggi Pilecap (m)`[j]-0.1)*2)*(temp$`Dimensi Pilecap P (m)`[j]/(temp$`Jarak Tulangan Bawah (mm)`[j]/1000)+1)+(temp$`Dimensi Pilecap P (m)`[j]-0.1+(temp$`Tinggi Pilecap (m)`[j]-0.1)*2)*(temp$`Dimensi Pilecap L (m)`[j]/(temp$`Jarak Tulangan Bawah (mm)`[j]/1000)+1))*1.05*temp$`Jumlah Pilecap (bh)`[j]/12
        temp[[col_name]][j] <- prev_value+new_value
      bbb_pilecap_step_2 <- temp
    }
    
    if(bbb_pilecap_step_2$`Diameter Sabuk (mm)`[j] == k){
      temp <- bbb_pilecap_step_2
      col_name <- paste0("D",k," - btg")
      prev_value <- temp[[col_name]][j]
      new_value <- (temp$`Dimensi Pilecap L (m)`[j]*2+temp$`Dimensi Pilecap P (m)`[j]*2)*temp$`Jumlah Sabuk (bh)`[j]*1.05/12*temp$`Jumlah Pilecap (bh)`[j]
      temp[[col_name]][j] <- prev_value+new_value
      bbb_pilecap_step_2 <- temp
    }
    
  }
  
}

for (k in diameter_unique){
  temp <- bbb_pilecap_step_2
  temp[[paste0("D",k," - kg")]] <- temp[[paste0("D",k," - btg")]] * massa_besi$`Berat (kg)`[massa_besi$`Diameter (mm)` == k]
  bbb_pilecap_step_2 <- temp
}

for (k in diameter_unique){
  
  temp <- bbb_pilecap_step_2
  
  col_name <- paste0("D",k," - btg")
  prev_value <- temp$`Total - btg`
  new_value <- temp[[col_name]] + prev_value
  temp$`Total - btg` <- new_value
  
  col_name <- paste0("D",k," - kg")
  prev_value <- temp$`Total - kg`
  new_value <- temp[[col_name]] + prev_value
  temp$`Total - kg` <- new_value
  
  bbb_pilecap_step_2 <- temp
  
}

bbb_pilecap_step_3 <- add_total_row(bbb_pilecap,bbb_pilecap_step_2,0)
bbb_pilecap_step_3$`CEK (kg/m3)` <- as.numeric(bbb_pilecap_step_3$`Total - kg`)/as.numeric(bbb_pilecap_step_3$`Volume Beton (m3)`)


#BBB - SLOOF
bbb_sloof_step_1 <- bbb_sloof %>%
  mutate(`Volume Beton (m3)` = `Panjang Sloof (m)`*`Dimensi Sloof B (m)`*`Dimensi Sloof H (m)`,
         `Begesting (m2)` = `Panjang Sloof (m)`*2*`Dimensi Sloof H (m)`)

bbb_sloof_step_2 <- d_table(bbb_sloof_step_1)

for (j in 1:nrow(bbb_sloof_step_2)){
  
  for (k in diameter_unique){
    
    if(bbb_sloof_step_2$`Diameter Tulangan Utama (mm)`[j] == k){
      temp <- bbb_sloof_step_2
      col_name <- paste0("D",k," - btg")
      prev_value <- temp[[col_name]][j]
      new_value <- temp$`Panjang Sloof (m)`[j]*(temp$`Jumlah Tulangan Tumpuan Atas (bh)`[j]+temp$`Jumlah Tulangan Tumpuan Bawah (bh)`[j]+temp$`Jumlah Tulangan Lapangan Atas (bh)`[j]+temp$`Jumlah Tulangan Lapangan Bawah (bh)`[j])/2*1.1/12
      temp[[col_name]][j] <- prev_value+new_value
      bbb_sloof_step_2 <- temp
    }
    
    if(bbb_sloof_step_2$`Diameter Begel (mm)`[j] == k){
      temp <- bbb_sloof_step_2
      col_name <- paste0("D",k," - btg")
      prev_value <- temp[[col_name]][j]
      new_value <- (temp$`Panjang Sloof (m)`[j]/(temp$`Jarak Begel Tumpuan (mm)`[j]/1000)*(temp$`Dimensi Sloof B (m)`[j]*2+temp$`Dimensi Sloof H (m)`[j]*2+temp$`Dimensi Sloof H (m)`[j]*temp$`Jumlah Begel I Tumpuan (bh)`[j])/2+temp$`Panjang Sloof (m)`[j]/(temp$`Jarak Begel Lapangan (mm)`[j]/1000)*(temp$`Dimensi Sloof B (m)`[j]*2+temp$`Dimensi Sloof H (m)`[j]*2+temp$`Dimensi Sloof H (m)`[j]*temp$`Jumlah Begel I Lapangan (bh)`[j])/2)*1.05/12
      temp[[col_name]][j] <- prev_value+new_value
      bbb_sloof_step_2 <- temp
    }
    
    if(bbb_sloof_step_2$`Diameter Sabuk (mm)`[j] == k){
      temp <- bbb_sloof_step_2
      col_name <- paste0("D",k," - btg")
      prev_value <- temp[[col_name]][j]
      new_value <- temp$`Panjang Sloof (m)`[j]*temp$`Jumlah Sabuk (bh)`[j]*1.05/12
      temp[[col_name]][j] <- prev_value+new_value
      bbb_sloof_step_2 <- temp
    }
    
  }
  
}

for (k in diameter_unique){
  temp <- bbb_sloof_step_2
  temp[[paste0("D",k," - kg")]] <- temp[[paste0("D",k," - btg")]] * massa_besi$`Berat (kg)`[massa_besi$`Diameter (mm)` == k]
  bbb_sloof_step_2 <- temp
}

for (k in diameter_unique){
  
  temp <- bbb_sloof_step_2
  
  col_name <- paste0("D",k," - btg")
  prev_value <- temp$`Total - btg`
  new_value <- temp[[col_name]] + prev_value
  temp$`Total - btg` <- new_value
  
  col_name <- paste0("D",k," - kg")
  prev_value <- temp$`Total - kg`
  new_value <- temp[[col_name]] + prev_value
  temp$`Total - kg` <- new_value
  
  bbb_sloof_step_2 <- temp
  
}

bbb_sloof_step_3 <- add_total_row(bbb_sloof,bbb_sloof_step_2,0)
bbb_sloof_step_3$`CEK (kg/m3)` <- as.numeric(bbb_sloof_step_3$`Total - kg`)/as.numeric(bbb_sloof_step_3$`Volume Beton (m3)`)

#BBB - KOLOM LANTAI
lantai <- unique(bbb_kolom_lantai$LANTAI)

for(l in lantai){
  
  bbb_kolom_lantai_n <- bbb_kolom_lantai[bbb_kolom_lantai$LANTAI == l,]
  bbb_kolom_lantai_n <- bbb_kolom_lantai_n[,-1]
  
  bbb_kolom_lantai_step_1 <- bbb_kolom_lantai_n %>%
    mutate(`Volume Beton (m3)` = `Jumlah Kolom (bh)`*`Panjang Kolom T (m)`*`Dimensi Kolom B (m)`*`Dimensi Kolom H (m)`,
           `Begesting (m2)` = `Jumlah Kolom (bh)`*`Panjang Kolom T (m)`*(`Dimensi Kolom B (m)`*2+`Dimensi Kolom H (m)`*2))
  
  bbb_kolom_lantai_step_2 <- d_table(bbb_kolom_lantai_step_1)
  
  for (j in 1:nrow(bbb_kolom_lantai_step_2)){
    
    for (k in diameter_unique){
      
      if(bbb_kolom_lantai_step_2$`Diameter Tulangan Utama (mm)`[j] == k){
        temp <- bbb_kolom_lantai_step_2
        col_name <- paste0("D",k," - btg")
        prev_value <- temp[[col_name]][j]
        if(l == 1){
          new_value <- (temp$`Panjang Kolom T (m)`[j]+40*temp$`Diameter Tulangan Utama (mm)`[j]/1000+0.3)*temp$`Jumlah Tulangan Utama`[j]*temp$`Jumlah Kolom (bh)`[j]/12*1.05
        }else{
          new_value <- (temp$`Panjang Kolom T (m)`[j]+40*temp$`Diameter Tulangan Utama (mm)`[j]/1000)*temp$`Jumlah Tulangan Utama`[j]*temp$`Jumlah Kolom (bh)`[j]/12*1.05
        }
        temp[[col_name]][j] <- prev_value+new_value
        bbb_kolom_lantai_step_2 <- temp
      }
      
      if(bbb_kolom_lantai_step_2$`Diameter Begel (mm)`[j] == k){
        temp <- bbb_kolom_lantai_step_2
        col_name <- paste0("D",k," - btg")
        prev_value <- temp[[col_name]][j]
        new_value <- (temp$`Panjang Kolom T (m)`[j]/(temp$`Jarak Begel Tumpuan (mm)`[j]/1000)*(temp$`Dimensi Kolom B (m)`[j]*2+temp$`Dimensi Kolom H (m)`[j]*2+temp$`Begel I Tumpuan (bh)`[j]*temp$`Dimensi Kolom H (m)`[j]+temp$`Begel I Tumpuan (bh)`[j]*temp$`Dimensi Kolom B (m)`[j])/2+temp$`Panjang Kolom T (m)`[j]/(temp$`Jarak Begel Lapangan (mm)`[j]/1000)*(temp$`Dimensi Kolom B (m)`[j]*2+temp$`Dimensi Kolom H (m)`[j]*2+temp$`Begel I Lapangan (bh)`[j]*temp$`Dimensi Kolom H (m)`[j]+temp$`Begel I Lapangan (bh)`[j]*temp$`Dimensi Kolom B (m)`[j])/2)*1.05/12*temp$`Jumlah Kolom (bh)`[j]
        temp[[col_name]][j] <- prev_value+new_value
        bbb_kolom_lantai_step_2 <- temp
      }
      
    }
    
  }
  
  for (k in diameter_unique){
    temp <- bbb_kolom_lantai_step_2
    temp[[paste0("D",k," - kg")]] <- temp[[paste0("D",k," - btg")]] * massa_besi$`Berat (kg)`[massa_besi$`Diameter (mm)` == k]
    bbb_kolom_lantai_step_2 <- temp
  }
  
  for (k in diameter_unique){
    
    temp <- bbb_kolom_lantai_step_2
    
    col_name <- paste0("D",k," - btg")
    prev_value <- temp$`Total - btg`
    new_value <- temp[[col_name]] + prev_value
    temp$`Total - btg` <- new_value
    
    col_name <- paste0("D",k," - kg")
    prev_value <- temp$`Total - kg`
    new_value <- temp[[col_name]] + prev_value
    temp$`Total - kg` <- new_value
    
    bbb_kolom_lantai_step_2 <- temp
    
  }
  
  bbb_kolom_lantai_step_3 <- add_total_row(bbb_kolom_lantai,bbb_kolom_lantai_step_2,1)
  bbb_kolom_lantai_step_3$`CEK1 (kg/m3)` <- as.numeric(bbb_kolom_lantai_step_3$`Total - kg`)/as.numeric(bbb_kolom_lantai_step_3$`Volume Beton (m3)`)
  bbb_kolom_lantai_step_3$`CEK2 (m2/m3)` <- as.numeric(bbb_kolom_lantai_step_3$`Begesting (m2)`)/as.numeric(bbb_kolom_lantai_step_3$`Volume Beton (m3)`)
  
  assign(paste0("bbb_kolom_lantai_",l,"_step_3"), bbb_kolom_lantai_step_3)
  
}

#BBB - PITLIFT
bbb_pitlift_step_1 <- bbb_pitlift %>%
  mutate(`Total Luas (m2)` = (`Panjang (m)`*2*(`Tebal Plat (mm)`/1000 +`Dalam Pitlift (mm)`/1000))+((`Lebar (m)`+`Tebal Plat (mm)`/1000)*2 *(`Tebal Plat (mm)`/1000 +`Dalam Pitlift (mm)`/1000)),
         `Volume Beton (m3)` = `Total Luas (m2)`*`Tebal Plat (mm)`/1000,
         `Begesting (m2)` = `Total Luas (m2)`*2)

bbb_pitlift_step_2 <- d_table(bbb_pitlift_step_1)

for (j in 1:nrow(bbb_pitlift_step_2)){
  
  for (k in diameter_unique){
    
    if(bbb_pitlift_step_2$`Diameter Besi X (mm)`[j] == k){
      temp <- bbb_pitlift_step_2
      col_name <- paste0("D",k," - btg")
      prev_value <- temp[[col_name]][j]
      new_value <- (temp$`Panjang (m)`[j]*2+(temp$`Lebar (m)`[j]+temp$`Tebal Plat (mm)`[j]/1000)*2)/temp$`Jarak Besi X (mm)`[j]*1000*2*1.1*((temp$`Tebal Plat (mm)`[j]/1000+temp$`Dalam Pitlift (mm)`[j]/1000)+(temp$`Tebal Pilecap (mm)`[j]/1000))/12
      temp[[col_name]][j] <- prev_value+new_value
      bbb_pitlift_step_2 <- temp
    }
    
    if(bbb_pitlift_step_2$`Diameter Besi Y(mm)`[j] == k){
      temp <- bbb_pitlift_step_2
      col_name <- paste0("D",k," - btg")
      prev_value <- temp[[col_name]][j]
      new_value <- temp$`Total Luas (m2)`[j]/temp$`Jarak Y(mm)`[j]*1000*2*1.15/12 
      temp[[col_name]][j] <- prev_value+new_value
      bbb_pitlift_step_2 <- temp
    }
    
  }
  
}

for (k in diameter_unique){
  temp <- bbb_pitlift_step_2
  temp[[paste0("D",k," - kg")]] <- temp[[paste0("D",k," - btg")]] * massa_besi$`Berat (kg)`[massa_besi$`Diameter (mm)` == k]
  bbb_pitlift_step_2 <- temp
}

for (k in diameter_unique){
  
  temp <- bbb_pitlift_step_2
  
  col_name <- paste0("D",k," - btg")
  prev_value <- temp$`Total - btg`
  new_value <- temp[[col_name]] + prev_value
  temp$`Total - btg` <- new_value
  
  col_name <- paste0("D",k," - kg")
  prev_value <- temp$`Total - kg`
  new_value <- temp[[col_name]] + prev_value
  temp$`Total - kg` <- new_value
  
  bbb_pitlift_step_2 <- temp
  
}

bbb_pitlift_step_2$`CEK1 (kg/m3)` <- as.numeric(bbb_pitlift_step_2$`Total - kg`)/as.numeric(bbb_pitlift_step_2$`Volume Beton (m3)`)
bbb_pitlift_step_2$`CEK2 (m2/m3)` <- as.numeric(bbb_pitlift_step_2$`Begesting (m2)`)/as.numeric(bbb_pitlift_step_2$`Volume Beton (m3)`)

#BBB - KOLOM LIFT
bbb_kolom_lift_step_1 <- bbb_kolom_lift %>%
  mutate(`Volume Beton (m3)` = `Jumlah Kolom (bh)`*`Panjang Kolom T (m)`*`Dimensi Kolom B (m)`*`Dimensi Kolom H (m)`,
         `Begesting (m2)` = `Jumlah Kolom (bh)`*`Panjang Kolom T (m)`*(`Dimensi Kolom B (m)`*2+`Dimensi Kolom H (m)`*2))

bbb_kolom_lift_step_2 <- d_table(bbb_kolom_lift_step_1)

for (j in 1:nrow(bbb_kolom_lift_step_2)){
  
  for (k in diameter_unique){
    
    if(bbb_kolom_lift_step_2$`Diameter Tulangan Utama (mm)`[j] == k){
      temp <- bbb_kolom_lift_step_2
      col_name <- paste0("D",k," - btg")
      prev_value <- temp[[col_name]][j]
      new_value <- (temp$`Panjang Kolom T (m)`[j]+40*temp$`Diameter Tulangan Utama (mm)`[j]/1000+0.3)*temp$`Jumlah Tulangan Utama`[j]*temp$`Jumlah Kolom (bh)`[j]/12*1.05
      temp[[col_name]][j] <- prev_value+new_value
      bbb_kolom_lift_step_2 <- temp
    }
    
    if(bbb_kolom_lift_step_2$`Diameter Begel (mm)`[j] == k){
      temp <- bbb_kolom_lift_step_2
      col_name <- paste0("D",k," - btg")
      prev_value <- temp[[col_name]][j]
      new_value <- (temp$`Panjang Kolom T (m)`[j]/(temp$`Jarak Begel Tumpuan (mm)`[j]/1000)*(temp$`Dimensi Kolom B (m)`[j]*2+temp$`Dimensi Kolom H (m)`[j]*2+temp$`Begel I Tumpuan (bh)`[j]*temp$`Dimensi Kolom H (m)`[j]+temp$`Begel I Tumpuan (bh)`[j]*temp$`Dimensi Kolom B (m)`[j])/2+temp$`Panjang Kolom T (m)`[j]/(temp$`Jarak Begel Lapangan (mm)`[j]/1000)*(temp$`Dimensi Kolom B (m)`[j]*2+temp$`Dimensi Kolom H (m)`[j]*2+temp$`Begel I Lapangan (bh)`[j]*temp$`Dimensi Kolom H (m)`[j]+temp$`Begel I Lapangan (bh)`[j]*temp$`Dimensi Kolom B (m)`[j])/2)*1.05/12*temp$`Jumlah Kolom (bh)`[j]
      temp[[col_name]][j] <- prev_value+new_value
      bbb_kolom_lift_step_2 <- temp
    }
    
  }
  
}

for (k in diameter_unique){
  temp <- bbb_kolom_lift_step_2
  temp[[paste0("D",k," - kg")]] <- temp[[paste0("D",k," - btg")]] * massa_besi$`Berat (kg)`[massa_besi$`Diameter (mm)` == k]
  bbb_kolom_lift_step_2 <- temp
}

for (k in diameter_unique){
  
  temp <- bbb_kolom_lift_step_2
  
  col_name <- paste0("D",k," - btg")
  prev_value <- temp$`Total - btg`
  new_value <- temp[[col_name]] + prev_value
  temp$`Total - btg` <- new_value
  
  col_name <- paste0("D",k," - kg")
  prev_value <- temp$`Total - kg`
  new_value <- temp[[col_name]] + prev_value
  temp$`Total - kg` <- new_value
  
  bbb_kolom_lift_step_2 <- temp
  
}

bbb_kolom_lift_step_3 <- add_total_row(bbb_kolom_lift,bbb_kolom_lift_step_2,0)
bbb_kolom_lift_step_3$`CEK1 (kg/m3)` <- as.numeric(bbb_kolom_lift_step_3$`Total - kg`)/as.numeric(bbb_kolom_lift_step_3$`Volume Beton (m3)`)
bbb_kolom_lift_step_3$`CEK2 (m2/m3)` <- as.numeric(bbb_kolom_lift_step_3$`Begesting (m2)`)/as.numeric(bbb_kolom_lift_step_3$`Volume Beton (m3)`)

#BBB - BALOK LIFT
bbb_balok_lift_step_1 <- bbb_balok_lift %>%
  mutate(`Volume Beton (m3)` = `Panjang Balok (m)`*`Dimensi Balok B (m)`*`Dimensi Balok H (m)`,
         `Begesting (m2)` = `Panjang Balok (m)`*`Dimensi Balok H (m)`*2)

bbb_balok_lift_step_2 <- d_table(bbb_balok_lift_step_1)

for (j in 1:nrow(bbb_balok_lift_step_2)){
  
  for (k in diameter_unique){
    
    if(bbb_balok_lift_step_2$`Diameter Tulangan Utama (mm)`[j] == k){
      temp <- bbb_balok_lift_step_2
      col_name <- paste0("D",k," - btg")
      prev_value <- temp[[col_name]][j]
      new_value <- temp$`Panjang Balok (m)`[j]*(temp$`Jumlah Tulangan Tumpuan Atas (bh)`[j]+temp$`Jumlah Tulangan Tumpuan Bawah (bh)`[j]+temp$`Jumlah Tulangan Lapangan Atas (bh)`[j]+temp$`Jumlah Tulangan Lapangan Bawah (bh)`[j])/2*1.1/12
      temp[[col_name]][j] <- prev_value+new_value
      bbb_balok_lift_step_2 <- temp
    }
    
    if(bbb_balok_lift_step_2$`Diameter Begel (mm)`[j] == k){
      temp <- bbb_balok_lift_step_2
      col_name <- paste0("D",k," - btg")
      prev_value <- temp[[col_name]][j]
      new_value <- (temp$`Panjang Balok (m)`[j]/(temp$`Jarak Begel Tumpuan (mm)`[j]/1000)*(temp$`Dimensi Balok B (m)`[j]*2+temp$`Dimensi Balok H (m)`[j]*2+temp$`Dimensi Balok H (m)`[j]*temp$`Jumlah Begel I Tumpuan (bh)`[j])/2+temp$`Panjang Balok (m)`[j]/(temp$`Jarak Begel Lapangan (mm)`[j]/1000)*(temp$`Dimensi Balok B (m)`[j]*2+temp$`Dimensi Balok H (m)`[j]*2+temp$`Dimensi Balok H (m)`[j]*(temp$`Jumlah Begel I Lapangan (bh)`[j]))/2)*1.05/12
      temp[[col_name]][j] <- prev_value+new_value
      bbb_balok_lift_step_2 <- temp
    }
    
    if(bbb_balok_lift_step_2$`Diameter Sabuk (mm)`[j] == k){
      temp <- bbb_balok_lift_step_2
      col_name <- paste0("D",k," - btg")
      prev_value <- temp[[col_name]][j]
      new_value <- temp$`Panjang Balok (m)`[j]*temp$`Jumlah Sabuk (bh)`[j]*1.05/12
      temp[[col_name]][j] <- prev_value+new_value
      bbb_balok_lift_step_2 <- temp
    }
    
  }
  
}

for (k in diameter_unique){
  temp <- bbb_balok_lift_step_2
  temp[[paste0("D",k," - kg")]] <- temp[[paste0("D",k," - btg")]] * massa_besi$`Berat (kg)`[massa_besi$`Diameter (mm)` == k]
  bbb_balok_lift_step_2 <- temp
}

for (k in diameter_unique){
  
  temp <- bbb_balok_lift_step_2
  
  col_name <- paste0("D",k," - btg")
  prev_value <- temp$`Total - btg`
  new_value <- temp[[col_name]] + prev_value
  temp$`Total - btg` <- new_value
  
  col_name <- paste0("D",k," - kg")
  prev_value <- temp$`Total - kg`
  new_value <- temp[[col_name]] + prev_value
  temp$`Total - kg` <- new_value
  
  bbb_balok_lift_step_2 <- temp
  
}

bbb_balok_lift_step_3 <- add_total_row(bbb_balok_lift,bbb_balok_lift_step_2,0)
bbb_balok_lift_step_3$`CEK1 (kg/m3)` <- as.numeric(bbb_balok_lift_step_3$`Total - kg`)/as.numeric(bbb_balok_lift_step_3$`Volume Beton (m3)`)
bbb_balok_lift_step_3$`CEK2 (m2/m3)` <- as.numeric(bbb_balok_lift_step_3$`Begesting (m2)`)/as.numeric(bbb_balok_lift_step_3$`Volume Beton (m3)`)


#BBB - BALOK LANTAI
lantai <- unique(bbb_balok_lantai$LANTAI)

for(l in lantai){
  
  bbb_balok_lantai_n <- bbb_balok_lantai[bbb_balok_lantai$LANTAI == l,]
  bbb_balok_lantai_n <- bbb_balok_lantai_n[,-1]
  
  bbb_balok_lantai_step_1 <- bbb_balok_lantai_n %>%
    mutate(`Volume Beton (m3)` = `Panjang Balok (m)`*`Dimensi Balok B (m)`*`Dimensi Balok H (m)`,
           `Begesting (m2)` = `Panjang Balok (m)`*`Dimensi Balok H (m)`*2)
  
  bbb_balok_lantai_step_2 <- d_table(bbb_balok_lantai_step_1)
  
  for (j in 1:nrow(bbb_balok_lantai_step_2)){
    
    for (k in diameter_unique){
      
      if(bbb_balok_lantai_step_2$`Diameter Tulangan Utama (mm)`[j] == k){
        temp <- bbb_balok_lantai_step_2
        col_name <- paste0("D",k," - btg")
        prev_value <- temp[[col_name]][j]
        new_value <- temp$`Panjang Balok (m)`[j]*(temp$`Jumlah Tulangan Tumpuan Atas (bh)`[j]+temp$`Jumlah Tulangan Tumpuan Bawah (bh)`[j]+temp$`Jumlah Tulangan Lapangan Atas (bh)`[j]+temp$`Jumlah Tulangan Lapangan Bawah (bh)`[j])/2*1.1/12
        temp[[col_name]][j] <- prev_value+new_value
        bbb_balok_lantai_step_2 <- temp
      }
      
      if(bbb_balok_lantai_step_2$`Diameter Begel (mm)`[j] == k){
        temp <- bbb_balok_lantai_step_2
        col_name <- paste0("D",k," - btg")
        prev_value <- temp[[col_name]][j]
        new_value <- (temp$`Panjang Balok (m)`[j]/(temp$`Jarak Begel Tumpuan (mm)`[j]/1000)*(temp$`Dimensi Balok B (m)`[j]*2+temp$`Dimensi Balok H (m)`[j]*2+temp$`Dimensi Balok H (m)`[j]*temp$`Jumlah Begel I Tumpuan (bh)`[j])/2+temp$`Panjang Balok (m)`[j]/(temp$`Jarak Begel Lapangan (mm)`[j]/1000)*(temp$`Dimensi Balok B (m)`[j]*2+temp$`Dimensi Balok H (m)`[j]*2+temp$`Dimensi Balok H (m)`[j]*(temp$`Jumlah Begel I Lapangan (bh)`[j]))/2)*1.05/12
        temp[[col_name]][j] <- prev_value+new_value
        bbb_balok_lantai_step_2 <- temp
      }
      
      if(bbb_balok_lantai_step_2$`Diameter Sabuk (mm)`[j] == k){
        temp <- bbb_balok_lantai_step_2
        col_name <- paste0("D",k," - btg")
        prev_value <- temp[[col_name]][j]
        new_value <- temp$`Panjang Balok (m)`[j]*temp$`Jumlah Sabuk (bh)`[j]*1.05/12
        temp[[col_name]][j] <- prev_value+new_value
        bbb_balok_lantai_step_2 <- temp
      }
      
    }
    
  }
  
  for (k in diameter_unique){
    temp <- bbb_balok_lantai_step_2
    temp[[paste0("D",k," - kg")]] <- temp[[paste0("D",k," - btg")]] * massa_besi$`Berat (kg)`[massa_besi$`Diameter (mm)` == k]
    bbb_balok_lantai_step_2 <- temp
  }
  
  for (k in diameter_unique){
    
    temp <- bbb_balok_lantai_step_2
    
    col_name <- paste0("D",k," - btg")
    prev_value <- temp$`Total - btg`
    new_value <- temp[[col_name]] + prev_value
    temp$`Total - btg` <- new_value
    
    col_name <- paste0("D",k," - kg")
    prev_value <- temp$`Total - kg`
    new_value <- temp[[col_name]] + prev_value
    temp$`Total - kg` <- new_value
    
    bbb_balok_lantai_step_2 <- temp
    
  }
  
  bbb_balok_lantai_step_3 <- add_total_row(bbb_balok_lantai,bbb_balok_lantai_step_2,1)
  bbb_balok_lantai_step_3$`CEK1 (kg/m3)` <- as.numeric(bbb_balok_lantai_step_3$`Total - kg`)/as.numeric(bbb_balok_lantai_step_3$`Volume Beton (m3)`)
  bbb_balok_lantai_step_3$`CEK2 (m2/m3)` <- as.numeric(bbb_balok_lantai_step_3$`Begesting (m2)`)/as.numeric(bbb_balok_lantai_step_3$`Volume Beton (m3)`)
  
  assign(paste0("bbb_balok_lantai_",l,"_step_3"), bbb_balok_lantai_step_3)
  
}

#BBB - BOREPILE
bbb_borepile_step_1 <- bbb_borepile %>%
  mutate(`Volume Beton (m3)` = ((`Diameter Borepile D (m)`/2)^2)*3.14*`Jumlah Borepile (bh)`*`Panjang Borepile T (m)`)

bbb_borepile_step_2 <- d_table(bbb_borepile_step_1)

for (j in 1:nrow(bbb_borepile_step_2)){
  
  for (k in diameter_unique){
    
    if(bbb_borepile_step_2$`Diameter Tulangan Utama (mm)`[j] == k){
      temp <- bbb_borepile_step_2
      col_name <- paste0("D",k," - btg")
      prev_value <- temp[[col_name]][j]
      new_value <- (temp$`Panjang Borepile T (m)`[j]+40*temp$`Diameter Tulangan Utama (mm)`[j]/1000)*temp$`Jumlah Tulangan Utama`[j]*temp$`Jumlah Borepile (bh)`[j]/12*1.05
      temp[[col_name]][j] <- prev_value+new_value
      bbb_borepile_step_2 <- temp
    }
    
    if(bbb_borepile_step_2$`Diameter Begel (mm)`[j] == k){
      temp <- bbb_borepile_step_2
      col_name <- paste0("D",k," - btg")
      prev_value <- temp[[col_name]][j]
      new_value <- temp$`Panjang Borepile T (m)`[j]/(temp$`Jarak Begel (mm)`[j]/1000)*2*3.14*temp$`Diameter Borepile D (m)`[j]/2*temp$`Jumlah Borepile (bh)`[j]*1.05/12
      temp[[col_name]][j] <- prev_value+new_value
      bbb_borepile_step_2 <- temp
    }
    
  }
  
}

for (k in diameter_unique){
  temp <- bbb_borepile_step_2
  temp[[paste0("D",k," - kg")]] <- temp[[paste0("D",k," - btg")]] * massa_besi$`Berat (kg)`[massa_besi$`Diameter (mm)` == k]
  bbb_borepile_step_2 <- temp
}

for (k in diameter_unique){
  
  temp <- bbb_borepile_step_2
  
  col_name <- paste0("D",k," - btg")
  prev_value <- temp$`Total - btg`
  new_value <- temp[[col_name]] + prev_value
  temp$`Total - btg` <- new_value
  
  col_name <- paste0("D",k," - kg")
  prev_value <- temp$`Total - kg`
  new_value <- temp[[col_name]] + prev_value
  temp$`Total - kg` <- new_value
  
  bbb_borepile_step_2 <- temp
  
}

bbb_borepile_step_3 <- add_total_row(bbb_borepile,bbb_borepile_step_2,0)
bbb_borepile_step_3$`CEK1 (kg/m3)` <- as.numeric(bbb_borepile_step_3$`Total - kg`)/as.numeric(bbb_borepile_step_3$`Volume Beton (m3)`)

#BBB - TANGGA
lantai <- unique(bbb_tangga$LANTAI)

for(l in lantai){
  
  bbb_tangga_n <- bbb_tangga[bbb_tangga$LANTAI == l,]
  bbb_tangga_n <- bbb_tangga_n[,-1]
  
  bbb_tangga_step_1 <- bbb_tangga_n %>%
    mutate(
      `Volume Beton (m3)` = `Jumlah Tangga (bh)`*`Jumlah Anak Tangga (bh)`*`Lebar Tangga Lt (m)`*(`Tebal Tangga H (m)`+`Tebal Tangga H (m)`+0.18)/2*0.3+`Lebar Bordes Lb (m)`*`Tebal Tangga H (m)`*`Lebar Tangga Lt (m)`*2+0.2*0.4*`Lebar Tangga Lt (m)`*2+0.2*1.2*`Lebar Tangga Lt (m)`,
      `Begesting (m2)` = (0.35*(`Lebar Tangga Lt (m)`+`Tebal Tangga H (m)`*2+0.2*2)+0.2*`Lebar Tangga Lt (m)`)*`Jumlah Anak Tangga (bh)`*`Jumlah Tangga (bh)`+(`Lebar Bordes Lb (m)`+0.4+0.4)*`Lebar Tangga Lt (m)`*2*`Jumlah Tangga (bh)`*1.2*(`Lebar Tangga Lt (m)`+0.4)
           )
  
  bbb_tangga_step_2 <- d_table(bbb_tangga_step_1)
  
  for (j in 1:nrow(bbb_tangga_step_2)){
    
    for (k in diameter_unique){
      
      if(bbb_tangga_step_2$`Diameter Tulangan Utama (mm)`[j] == k){
        temp <- bbb_tangga_step_2
        col_name <- paste0("D",k," - btg")
        prev_value <- temp[[col_name]][j]
        new_value <- ((temp$`Jumlah Anak Tangga (bh)`[j]*0.35+temp$`Lebar Bordes Lb (m)`[j]*2+0.2*2+2*2)*(temp$`Lebar Tangga Lt (m)`[j]/temp$`Jarak Tulanganangan Utama (mm)`[j]+1)*2+(temp$`Lebar Tangga Lt (m)`[j]*2+0.2)*((temp$`Lebar Bordes Lb (m)`[j]/temp$`Jarak Tulanganangan Utama (mm)`[j]+1)*2+5))*temp$`Jumlah Tangga (bh)`[j] / 12 * 1.05
        temp[[col_name]][j] <- prev_value+new_value
        bbb_tangga_step_2 <- temp
      }
      
      if(bbb_tangga_step_2$`Diameter Tulangan Bagi (mm)`[j] == k){
        temp <- bbb_tangga_step_2
        col_name <- paste0("D",k," - btg")
        prev_value <- temp[[col_name]][j]
        new_value <- (((temp$`Jumlah Anak Tangga (bh)`[j]*0.35/temp$`Jarak Tulangan Bagi (mm)`[j]+1)*2+1*temp$`Jumlah Anak Tangga (bh)`[j])*temp$`Lebar Tangga Lt (m)`[j]+0.8*temp$`Jumlah Anak Tangga (bh)`[j]*(temp$`Lebar Tangga Lt (m)`[j]/temp$`Jarak Tulangan Bagi (mm)`[j]+1))*1.05/12
        temp[[col_name]][j] <- prev_value+new_value
        bbb_tangga_step_2 <- temp
      }
      
    }
    
  }
  
  for (k in diameter_unique){
    temp <- bbb_tangga_step_2
    temp[[paste0("D",k," - kg")]] <- temp[[paste0("D",k," - btg")]] * massa_besi$`Berat (kg)`[massa_besi$`Diameter (mm)` == k]
    bbb_tangga_step_2 <- temp
  }
  
  for (k in diameter_unique){
    
    temp <- bbb_tangga_step_2
    
    col_name <- paste0("D",k," - btg")
    prev_value <- temp$`Total - btg`
    new_value <- temp[[col_name]] + prev_value
    temp$`Total - btg` <- new_value
    
    col_name <- paste0("D",k," - kg")
    prev_value <- temp$`Total - kg`
    new_value <- temp[[col_name]] + prev_value
    temp$`Total - kg` <- new_value
    
    bbb_tangga_step_2 <- temp
    
  }
  
  bbb_tangga_step_3 <- add_total_row(bbb_tangga,bbb_tangga_step_2,1)
  bbb_tangga_step_3$`CEK1 (kg/m3)` <- as.numeric(bbb_tangga_step_3$`Total - kg`)/as.numeric(bbb_tangga_step_3$`Volume Beton (m3)`)
  bbb_tangga_step_3$`CEK2 (m2/m3)` <- as.numeric(bbb_tangga_step_3$`Begesting (m2)`)/as.numeric(bbb_tangga_step_3$`Volume Beton (m3)`)
  
  assign(paste0("bbb_tangga_",l,"_step_3"), bbb_tangga_step_3)
  
}


#BBB - PLAT
lantai <- unique(bbb_plat2$LANTAI)

for(l in lantai){
  
  bbb_plat_luas <- bbb_plat1[bbb_plat1$LANTAI == l,]
  bbb_plat_luas$`Luas (m2)` <- bbb_plat_luas$`Panjang (m)` * bbb_plat_luas$`Lebar (m)`
  bbb_plat_luas <- bbb_plat_luas[,-1]
  
  total_luas <- sum(bbb_plat_luas$`Luas (m2)`)
  
  bbb_plat2_n <- bbb_plat2[bbb_plat2$LANTAI == l,]
  bbb_plat2_n$`Luas (m2)` <- total_luas
  bbb_plat2_n <- bbb_plat2_n[,-1]
  
  bbb_plat2_step_1 <- bbb_plat2_n %>%
    mutate(
      `Volume Beton (m3)` = `Luas (m2)`*`Tebal Plat (mm)`/1000,
      `Begesting (m2)` = `Luas (m2)`
    )
  
  bbb_plat2_step_2 <- d_table(bbb_plat2_step_1)
  
  for (j in 1:nrow(bbb_plat2_step_2)){
    
    for (k in diameter_unique){
      
      if(bbb_plat2_step_2$`Diameter Besi (mm)`[j] == k){
        temp <- bbb_plat2_step_2
        col_name <- paste0("D",k," - btg")
        prev_value <- temp[[col_name]][j]
        new_value <- (temp$`Luas (m2)`[j]/temp$`Jarak Besi X (mm)`[j]*1000*2*1.15/12) + (temp$`Luas (m2)`[j]/temp$`Jarak Besi Y (mm)`[j]*1000*2*1.15/12)
        temp[[col_name]][j] <- prev_value+new_value
        bbb_plat2_step_2 <- temp
      }
      
    }
    
  }
  
  for (k in diameter_unique){
    temp <- bbb_plat2_step_2
    temp[[paste0("D",k," - kg")]] <- temp[[paste0("D",k," - btg")]] * massa_besi$`Berat (kg)`[massa_besi$`Diameter (mm)` == k]
    bbb_plat2_step_2 <- temp
  }
  
  for (k in diameter_unique){
    
    temp <- bbb_plat2_step_2
    
    col_name <- paste0("D",k," - btg")
    prev_value <- temp$`Total - btg`
    new_value <- temp[[col_name]] + prev_value
    temp$`Total - btg` <- new_value
    
    col_name <- paste0("D",k," - kg")
    prev_value <- temp$`Total - kg`
    new_value <- temp[[col_name]] + prev_value
    temp$`Total - kg` <- new_value
    
    bbb_plat2_step_2 <- temp
    
  }
  
  #bbb_plat2_step_3 <- add_total_row(bbb_plat2,bbb_plat2_step_2)
  bbb_plat2_step_3 <- bbb_plat2_step_2
  bbb_plat2_step_3$`CEK1 (kg/m3)` <- as.numeric(bbb_plat2_step_3$`Total - kg`)/as.numeric(bbb_plat2_step_3$`Volume Beton (m3)`)
  bbb_plat2_step_3$`CEK2 (m2/m3)` <- as.numeric(bbb_plat2_step_3$`Begesting (m2)`)/as.numeric(bbb_plat2_step_3$`Volume Beton (m3)`)
  
  assign(paste0("bbb_plat_luas_",l), bbb_plat_luas)
  assign(paste0("bbb_plat2_",l,"_step_3"), bbb_plat2_step_3)
  
}

#PRINT FUNCTION

wb <- createWorkbook()

row_loc <- 1

addWorksheet(wb, sheetName = "PERHITUNGAN")

tulis <- function(x, row, col){
  writeData(wb, sheet = "PERHITUNGAN", x, startCol = col, startRow = row)
}

num_to_al <- function(n){
  result <- ""
  while(n > 0){
    n <- n-1
    remainder <- n %% 26
    result <- paste0(intToUtf8(remainder + 65), result)
    n <- n %/% 26
  }
  return(result)
}

#GALIAN
tulis("GALIAN", row_loc, 1)
row_loc <- row_loc + 2

tulis("A.", row_loc, 1)
tulis("PILECAP", row_loc, 2)
row_loc <- row_loc + 1
tulis(galian_pilecap_step_2, row_loc, 2)
row_loc <- row_loc + nrow(galian_pilecap_step_2) + 2

tulis("B.", row_loc, 1)
tulis("SLOOF", row_loc, 2)
row_loc <- row_loc + 1
tulis(galian_sloof_step_2, row_loc, 2)
row_loc <- row_loc + nrow(galian_sloof_step_2) + 3

#BBB
tulis("BESI, BETON, DAN BEGESTING", row_loc, 1)
row_loc <- row_loc + 2

tulis("A.1.", row_loc, 1)
tulis("PILECAP", row_loc, 2)
row_loc <- row_loc + 1
tulis(bbb_pilecap_step_3, row_loc, 2)
row_loc <- row_loc + nrow(bbb_pilecap_step_3) + 2

tulis("B.1.", row_loc, 1)
tulis("SLOOF", row_loc, 2)
row_loc <- row_loc + 1
tulis(bbb_sloof_step_3, row_loc, 2)
row_loc <- row_loc + nrow(bbb_sloof_step_3) + 2

ind <- 0
for(l in unique(bbb_kolom_lantai$LANTAI)){
  ind <- ind + 1
  tulis(paste0("C.",ind,"."),row_loc,1)
  tulis(paste0("KOLOM LANTAI ",l),row_loc,2)
  row_loc <- row_loc + 1
  temp_df <- get(paste0("bbb_kolom_lantai_",l,"_step_3"))
  tulis(temp_df,row_loc, 2)
  row_loc <- row_loc + nrow(temp_df) + 2
}

tulis("D.1.", row_loc, 1)
tulis("PITLIFT", row_loc, 2)
row_loc <- row_loc + 1
tulis(bbb_pitlift_step_2, row_loc, 2)
row_loc <- row_loc + nrow(bbb_pitlift_step_2) + 2

tulis("E.1.", row_loc, 1)
tulis("KOLOM LIFT", row_loc, 2)
row_loc <- row_loc + 1
tulis(bbb_kolom_lift_step_3, row_loc, 2)
row_loc <- row_loc + nrow(bbb_kolom_lift_step_3) + 2

tulis("F.1.", row_loc, 1)
tulis("BALOK LIFT", row_loc, 2)
row_loc <- row_loc + 1
tulis(bbb_balok_lift_step_3, row_loc, 2)
row_loc <- row_loc + nrow(bbb_balok_lift_step_3) + 2

ind <- 0
for(l in unique(bbb_balok_lantai$LANTAI)){
  ind <- ind + 1
  tulis(paste0("G.",ind,"."),row_loc,1)
  tulis(paste0("BALOK LANTAI ",l),row_loc,2)
  row_loc <- row_loc + 1
  temp_df <- get(paste0("bbb_balok_lantai_",l,"_step_3"))
  tulis(temp_df,row_loc, 2)
  row_loc <- row_loc + nrow(temp_df) + 2
}

tulis("H.1.", row_loc, 1)
tulis("BOREPILE", row_loc, 2)
row_loc <- row_loc + 1
tulis(bbb_borepile_step_3, row_loc, 2)
row_loc <- row_loc + nrow(bbb_borepile_step_3) + 2

ind <- 0
for(l in unique(bbb_tangga$LANTAI)){
  ind <- ind + 1
  tulis(paste0("I.",ind,"."),row_loc,1)
  tulis(paste0("TANGGA LANTAI ",l),row_loc,2)
  row_loc <- row_loc + 1
  temp_df <- get(paste0("bbb_tangga_",l,"_step_3"))
  tulis(temp_df,row_loc, 2)
  row_loc <- row_loc + nrow(temp_df) + 2
}

ind <- 0
for(l in unique(bbb_plat2$LANTAI)){
  ind <- ind + 1
  tulis(paste0("J.",ind,"."),row_loc,1)
  tulis(paste0("PLAT LANTAI ",l),row_loc,2)
  row_loc <- row_loc + 1
  
  temp <- get(paste0("bbb_plat_luas_",l))
  tulis(temp,row_loc,2)
  row_loc <- row_loc + nrow(temp) + 1
  
  temp_df <- get(paste0("bbb_plat2_",l,"_step_3"))
  tulis(temp_df,row_loc, 2)
  row_loc <- row_loc + nrow(temp_df) + 2
}

saveWorkbook(wb, "Output/PERHITUNGAN.xlsx", overwrite = TRUE)


