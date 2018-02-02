order.index = function(x){
    Order = order(x)
    col.index = ceiling(Order/nrow(x))
    row.index = Order - (col.index - 1)*nrow(x)
    return(cbind(row.index, col.index))
}
