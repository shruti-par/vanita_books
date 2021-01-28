library(readxl)
library(writexl)
sales <- read_xlsx("./files/Sales.xlsx")

final_data <- data.frame(Date = as.Date(character()), Receipt = as.numeric(), 
                         WebAmt = as.numeric(), Count = as.numeric(),
                         "Maanavi Jeevanatil Gudha Rahasya Part 1" = as.numeric(),	
                         "Maanavi Jeevanatil Gudha Rahasya Part 2" = as.numeric(),
                         "Maanavi Jeevanatil Gudha Rahasya Part 3" = as.numeric(),
                         "Maanavi Jeevanatil Gudha Rahasya Part 4" = as.numeric(),
                         "Maanavi Jeevanatil Gudha Rahasya Part 5" = as.numeric(),
                         "Maanavi Jeevanatil Gudha Rahasya Part 6" = as.numeric(),
                         "Maanavi Jeevanatil Gudha Rahasya Part 7" = as.numeric(),
                         "Aatmasiddhi" = as.numeric(),
                         "Agamyavani" = as.numeric(),
                         "Bharatiya Gudha Vidhya" = as.numeric(),
                         "Amaratvakade Vaatchaal" = as.numeric(),
                         "Jeevan Mukticha Marg" = as.numeric(),
                         "Shri Datta Durga Samvad" = as.numeric(),
                         "Ishwar Praptiche Marg" = as.numeric(),
                         "Ishwar Bhaktiche Anubhav" = as.numeric(),
                         "Vihangam Marg" = as.numeric(),
                         "Shri Datta Upanishad" = as.numeric(),
                         "Prophecies" = as.numeric(),
                         "The Path Divine" = as.numeric(),
                         "Spiritual Journey" = as.numeric(),
                         "Siddhyogyachya Sahavasaat - Part 1" = as.numeric(),
                         "Narmada" = as.numeric(),
                         "Narmada vahanmarg" = as.numeric(),
                         "Siddhyogyachya Sahavasaat - Part 2" = as.numeric(),
                         "Shri Swami Samarth Saptashati" = as.numeric(),
                         "Gurucharitra" = as.numeric(),
                         "Sankshipta Shri Gurucharitra" = as.numeric(),
                         "Shri Dattaleelamrut-Shri Siddhaleelamrut" = as.numeric(),
                         "Shri Durga Saptashati" = as.numeric(),
                         "Shri Durga Trishati" = as.numeric(),
                         "Shri Datta Gita" = as.numeric(),
                         "Paramanand Lahri" = as.numeric(),
                         "Mahayogi" = as.numeric(),
                         "Datta Vijay" = as.numeric(), 
                         check.names = FALSE, 
                         stringsAsFactors = FALSE)

for(i in 1:nrow(sales))
{
  final_data[i, "Date"] <- as.Date.character(sales$Date[i])
  final_data$Receipt[i] <- as.numeric(unlist(strsplit(sales$`Order #`[i], "_"))[2])
  final_data$WebAmt[i] <- sales$`N. Revenue`[i]
  final_data$Count[i] <- sales$`Items Sold`[i]
  lst <- lapply(unlist(strsplit(sales$`Product(s)`[i], ", ")), FUN = function(x) unlist(strsplit(x, "× ")))
  for(j in 1:length(lst))
  {
    sublst = unlist(lst[[j]])
    if(sublst[2] == "Maanavi Jeevanatil Gudha Rahasya - Set")
    {
      final_data[i, paste("Maanavi Jeevanatil Gudha Rahasya Part", 1:7)] <- as.numeric(sublst[1])
    }
    else if(sublst[2] == "Siddhyogyachya Sahavasaat Set of 2")
    {
      final_data[i, paste("Siddhyogyachya Sahavasaat - Part", 1:2)] <- as.numeric(sublst[1])
    }
    else
    {
      final_data[i, sublst[2]] = as.numeric(sublst[1])
    }
  }
  
}

final_data[is.na(final_data)] <- ""
final_data <- final_data[order(final_data$Receipt),]
# write.csv(final_data, "SalesWB_edited.csv", row.names = FALSE)
write_xlsx(final_data, "SalesWB_edited.xlsx")
