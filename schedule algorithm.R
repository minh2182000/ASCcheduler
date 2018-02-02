rm(list=ls());
if (!"xlsx" %in% rownames(installed.packages())){install.packages("xlsx")}
require(xlsx)
# setwd("_____")

#-----READ INPUT FROM EXCEL----------------------
# availability
  A = read.xlsx("Lab.xls", sheetIndex = 1, startRow = 2)
  tutors.names = as.character(A[,1]); A[,1] = NULL
  A = t(as.matrix(A)); rownames(A) = NULL
  # 0 = no, 1 = yes, 2 = if need be
  
# total hours (sheet 2)
  Tau = read.xlsx("Lab.xls", sheetIndex = 2)
  
# Range of number of tutor (sheet 3)
  Range = read.xlsx("Lab.xls", sheetIndex = 3, colIndex = 3:4); Range = as.matrix(Range)
  a = as.matrix(Range[,1]); b = as.matrix(Range[,2]); rm(Range)
  
#-----ALGORITHM------------------------------------
X = matrix(0, nrow = nrow(A), ncol = ncol(A)) # output matrix
  
#1. weight of if need be-----
pi = sum(A==1)/(sum(A==1)+sum(A==2))
A[which(A==2)] = pi

#2. Assign all scarce time ----
LAMDA = rowSums((A!=0)*1)
scarce.time = which(LAMDA <= a)
X[scarce.time,] = (A[scarce.time,] !=0)*1

other.time = which(LAMDA>a)

#3. iterative simulation to assign other time-----------
X0 = X # save original schedule in case of failure
max.assigned = 0
for (trial in 1:1000){ # repeat simulation until find a solution
  cat("trial #",trial, " \n")
  X = X0; assigned = assigned.1 = sum(X)
    while(sum(Tau-T)>0){
      assigned.1 = assigned
      #3.1. rows to run on this round---------
        running.time = other.time[which(rowSums(X[other.time,])<b[other.time,])]
        need.gap = pmax(0,a[running.time] - rowSums(X[running.time,]))
        peak.time    = running.time[which(need.gap == max(need.gap))]
        order = sample(peak.time) #shuffle the rows
      #3.2. run through each row to assign tutor---------
        for (i in order){
            T = colSums(X)
          #a. calculate likelihood score omega
            omega = A[i,]*(Tau - T)^2*(1-X[i,])
            if (max(omega)!=0){
          #b. select a tutor using probability on weight of omega
              j = sample(1:ncol(X), 1, prob = omega)
              X[i,j] = 1
              assigned = assigned + 1
            }
        }
      #3.3. stop iteration if making no progress  ---------
        if (assigned == assigned.1){cat("failed, assigned total of ",assigned, "\n")
                                # save best attempt:
                                    if (assigned > max.assigned){max.assigned = assigned; best.X = X}
                                    break
        }else{
            assigned.1 = assigned; T = colSums(X)}
    } 
  if (assigned == sum(Tau)){cat("Success! max.assigned = ", assigned, "\n"); best.X = X; break
  }else{cat("max.assigned = ", max.assigned, "\n")}
  
  if (sum(X) == sum(b)){cat("center's maximum positions reached \n"); break}
}

#-- OUTPUT ----------------------------------
#chedule conflict check
check = sum(((A-1)*X == -1))
if (check != 0){stop("check failed, schedule conflict")}else{
  print("check passed")
  
# reformat output
  X = best.X
  X = t(X); rownames(X) = tutors.names
  out.schedule = matrix(NA, ncol = ncol(X), nrow = nrow(X))
  for (j in 1:ncol(X)){
    tutors.working = which(X[,j] == 1)
    out.schedule[,j] = c(tutors.names[tutors.working], rep(NA, nrow(X) - length(tutors.working)))
  }
# output
  write.xlsx(out.schedule,"schedule_output.xlsx")
}
