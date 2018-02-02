rm(list = ls())
if (!"XLConnect" %in% rownames(installed.packages())){install.packages("XLConnect")}
require(XLConnect)
source("order.index.R")
#-- Running parameters (directory) ------------------------------------------
setwd("E:/schedule algorithm")
#-- READ EXCEL INPUTS---------------------------------------------------------
  wb = loadWorkbook("sample input.xlsx")
  
# Day list
  Days = readWorksheet(wb, sheet = 1, startRow = 1, endRow = 1, startCol = 2, header = FALSE, useCachedValues = TRUE)
  Days = unlist(Days)
  
# Availability
  Avail = readWorksheet(wb, sheet = 1, startCol = 2, startRow = 3, header = F, useCachedValues = TRUE)
  Avail = as.matrix(Avail)
  Avail[Avail == "(OK)"] = 2
  Avail[Avail == "OK"] = 1
  Avail[is.na(Avail)] = 0
  Avail = apply(Avail, 2, as.numeric)
  
# tutors Capabilities
  Caps = readWorksheet(wb, sheet = 2, useCachedValues = TRUE)
  rownames(Caps) = Caps[,1]; Caps[,1] = NULL
  Caps = as.matrix(Caps)
    
# Tutors' hour limit
  PMax = unlist(readWorksheet(wb, sheet = 3, useCachedValues = TRUE))
  
# Centers' limits
  limits = readWorksheet(wb, sheet = 4, startRow = 3, startCol = 3, header = F, useCachedValues = TRUE)
  L = as.matrix(limits)
  

#-- ALGORITHM------------------------------------------------------------------
  # 1. Add gap between days 
  gaps = which(Days != c(Days[length(Days)], Days[-length(Days)]))
  for (day in rev(gaps)){
    if (day == 1){Avail = cbind(0, Avail)
                  L = rbind(0, L)
    }else {
                  Avail = cbind(Avail[,1:(day - 1)],
                                    0,
                                    Avail[,day:ncol(Avail)]
                                    )
                  L = rbind(L[1:(day - 1),],
                            0,
                            L[day:nrow(L), ]
                            )
    }
  }
  Avail = cbind(Avail, 0)
  L = rbind(L, 0)
  
  # generate 3D availability array
  A = array(dim = c(nrow(Avail), ncol(Avail), ncol(Caps)))
  for (c in 1:ncol(Caps)){
    A[,,c] = diag(Caps[,c]) %*% Avail
  }
  
  #0. weight of if-need-be ----------------------
  A[which(A==2)] = 0.1
  
  #1. order A.jk ascending
  Order = order.index(apply(A, MARGIN = c(2,3), FUN = sum))
  #2. Iterate

  best.sumX = 0
  for (sim in 1:20){
    cat("\14"); cat(sim*2, "%", "     ", best.sumX, "/", sum(L))
    X = array(0, dim = dim(A))
    #2.1. Loop through the j,k order
    for (row in 1:nrow(Order)){
      j = Order[row, 1]; k = Order[row, 2]
      if (j == 1 | j == dim(A)[2]){next}
      omega = A[,j,k] /
        (1 + apply(A, MARGIN = c(1, 2), sum)[, j]) * 
        (PMax - apply(X, MARGIN = 1, sum)) * 
        (1 - apply(X, MARGIN = c(1, 2), sum)[, j]) * 
        (1 + apply(A, MARGIN = c(1, 2), sum)[, j + 1])^50 * 
        (1 + apply(A, MARGIN = c(1, 2), sum)[, j - 1])^50
    if (sum(omega) == 0){next}  
    #2.2 Sample tutor
    pick = sample(c(0, which(omega > 0)), min(L[j,k], sum(omega > 0)), 
                  prob = c(0, omega[omega > 0]))
    X[pick, j, k] = 1 # assign
    }
    
    #2.3 remove 1hr shifts
    for (k in 1:dim(X)[3]){
      indi.X = X[,,k]
      `indi.X-1` = cbind(0, indi.X[,-ncol(indi.X)])
      `indi.X+1` = cbind(indi.X[,-1] , 0)
      X.1h = (indi.X==1) * (`indi.X-1`==0) * (`indi.X+1`==0)
      X[,,k] = X[,,k] - X.1h
    }
    if (sum(X) > best.sumX){best.X = X; best.sumX = sum(X)}
  }
  
  #3. remove gaps
  for (day in gaps){
    print(sum(best.X[,day,]))
    best.X = best.X[,-day,]
  }
  best.X = best.X[,-dim(best.X)[2],]
#-- OUTPUT -------------------------------------------------------------------
  cat("\n outputting........................................... \n")
  rownames(best.X) = tutors.names = rownames(Caps)
  Times = readWorksheet(wb, sheet = 1, startRow = 1, endRow = 2, startCol = 2, header = FALSE)
  Times = paste(unlist(Times[1,]), unlist(Times[2,]))
  unlink("schedule output.xlsx")
  outwb = loadWorkbook("schedule output.xlsx", create = TRUE)
  #1. Centers' schedule---------------------------
  for (c in 1:ncol(Caps)){
    schedule =  best.X[,,c]
    out.schedule = matrix("", nrow = nrow(schedule), ncol = ncol(schedule))
    colnames(out.schedule) = Times
    for (j in 1:ncol(schedule)){
      tutors.working = which(schedule[,j] == 1)
      out.schedule[,j] = c(tutors.names[tutors.working], rep("", nrow(out.schedule) - length(tutors.working)))
    }
    
    createSheet(outwb, colnames(Caps)[c])
    writeWorksheet(outwb, data =  out.schedule, sheet = c)
  }
  #2. Individual hours-----------------------------
  createSheet(outwb, "Tutor Schedule")
  T.Schedule = as.matrix(apply(best.X, c(1), sum))
  writeWorksheet(outwb, data.frame(tutors.names, "Scheduled" = T.Schedule, "Max" = PMax, Caps), sheet = "Tutor Schedule")
  #3. save output ------------------------------------
  saveWorkbook(outwb)
  cat("done \n")
  
  