#' @name cellColor
#' 
#' @title cellColor - helper function for metIDQ_get_high_quality_features
cellColor <- function(style) {
    fg  <- style$getFillForegroundXSSFColor()
    rgb <- tryCatch(fg$getRgb(), error = function(e) "")
    rgb <- paste(rgb, collapse = "")
    return(rgb)
}

#' @name metIDQ_get_high_quality_features
#'
#' @title Get a vector with the high-quality features to retain
metIDQ_get_high_quality_features <- function(file, threshold = 0.66) {
    ## get the background color
    wb <- xlsx::loadWorkbook(file = file)
    sheet1 <- xlsx::getSheets(wb)[[1]]
    rows <- xlsx::getRows(sheet1)
    cells <- xlsx::getCells(rows)
    styles <- sapply(cells, xlsx::getCellStyle)
    bg <- sapply(styles, cellColor) 
    
    ## use values to get the indices where the actual data is stored
    values <- sapply(cells, xlsx::getCellValue)
    
    ## convert bg and values to matrix that it has the same dimension as before
    if (all(names(bg) != names(values))) 
        stop("dimensions do not match")
    row_index <- as.numeric(unlist(lapply(strsplit(names(bg), split = "[.]"), "[", 1)))
    col_index <- as.numeric(unlist(lapply(strsplit(names(bg), split = "[.]"), "[", 2)))
    mat_values <- mat_bg <- matrix("", ncol = max(col_index), nrow = max(row_index))
    for (i in 1:length(bg)) {
        mat_bg[row_index[i], col_index[i]] <- bg[i]
        mat_values[row_index[i], col_index[i]] <- values[i]
    }
    
    ## obtain the indices of the cells where the values/background colors are 
    ## stored in
    row_inds <- which(!is.na(mat_values[, 1]))
    row_inds <- row_inds[!row_inds %in% 1:18]
    col_inds <- rbind(which(mat_values == "C0", arr.ind = TRUE),
                      which(mat_values == "Choline", arr.ind = TRUE))
    colnames(mat_values) <- colnames(mat_bg) <- mat_values[col_inds[1, "row"], ]
    col_inds <- seq(col_inds[1, "col"], col_inds[2, "col"])
    
    ## iterate through the columns of mat_bg and check the color values 
    ## against the QC of MetIDQ## "00cd66" == green, "87ceeb" == lightblue
    valid <- apply(mat_bg[row_inds, col_inds], 2, 
                   function(x) sum(x %in% c("00cd66"))) 
    
    ## require that at least threshold*100% values per metabolite are 
    ## "green"/"lightblue", 
    features <- valid / length(row_inds)  > threshold
    
    ## only return the names of valid features
    make.names(names(which(features)))
}

#' @name add_vertex_attributes
#' 
#' @title Add vertex attributes to \code{igraph} object
#' 
#' @description 
#' The function adds vertex attributes to a \code{igraph} object \code{g}. The 
#' values are stored in the \code{attributes} object. The function will return
#' a \code{igraph} object with updated vertex attributes. 
#' 
#' The \code{attributes} object is a \code{data.frame}. 
#' 
#' The \code{data.frame} contains the columns \code{col_vertex}, a 
#' \code{character} vector of length 1, specifying the vertices and
#' the vertex attributes weights in the remaining columns. The
#' attributes will be stored in the respective slots with same names as the 
#' \code{colnames} of \code{attributes} for each vertex 
#' of the returned \code{igraph} object.
#' 
#' @details
#' \code{col_vertex} has to be adjusted such that it specifies the vertices. 
#' The \code{character} of length 1 will specify the column
#' containing the vertices of the graph.
#' 
#' @param g \code{igraph} object
#' @param attributes \code{data.frame} containing vertex
#' attribute information
#' @param col_vertex \code{chararacter} of length 1, specifying the column 
#' containing  vertices in \code{attributes} of type \code{data.frame}
#' 
#' @export
#' 
#' @return igraph object
#' 
#' @author Thomas Naake, \email{thomasnaake@@googlemail.com}
#'
#' @importFrom igraph set_vertex_attr V
#' 
#' @examples
#' FA <- c("FA(12:0)", "FA(14:0)", "FA(16:0)")
#' 
#' ## create data.frame with reactions and reaction order
#' reactions <- rbind(
#'     c(1, "RHEA:15421", "M_ATP + M_CoA + M_FA = M_PPi + M_AMP + M_AcylCoA", FALSE),
#'     c(2, "RHEA:15325", "M_Glycerol-3-P + M_AcylCoA = M_CoA + M_LPA", FALSE),
#'     c(3, "RHEA:19709", "M_AcylCoA + M_LPA = M_CoA + M_PA", FALSE),
#'     c(4, "RHEA:27429", "M_H2O + M_PA = M_Pi + M_1,2-DG", FALSE)
#' )
#' reactions <- data.frame(order = reactions[, 1], reaction_RHEA = reactions[, 2],
#'     reaction_formula = reactions[, 3], directed = reactions[, 4])
#' reactions$order <- as.numeric(reactions$order)
#' reactions$directed <- as.logical(reactions$directed)
#' 
#' ## create the list with reactions
#' reaction_l <- create_reactions(substrates = list(FA = FA), reactions = reactions)
#' 
#' ## create the adjacency matrix
#' reaction_adj <- create_reaction_adjacency_matrix(reaction_l)
#' 
#' ## create graph
#' g <- igraph::graph_from_adjacency_matrix(reaction_adj, mode = "directed", weighted = TRUE)
#' 
#' ## attributes: data.frame
#' attributes_df <- data.frame(
#'    name = c("CoA(12:0)", "CoA(14:0)", "CoA(16:0)", "DG(12:0/12:0/0:0)",
#'        "DG(12:0/14:0/0:0)", "DG(12:0/16:0/0:0)", "DG(14:0/12:0/0:0)",
#'        "DG(14:0/14:0/0:0)", "DG(14:0/16:0/0:0)", "DG(16:0/12:0/0:0)",
#'        "DG(16:0/14:0/0:0)", "DG(16:0/16:0/0:0)", "FA(12:0)", "FA(14:0)",
#'        "FA(16:0)", "PA(12:0/0:0)", "PA(12:0/12:0)", "PA(12:0/14:0)",
#'        "PA(12:0/16:0)", "PA(14:0/0:0)", "PA(14:0/12:0)", "PA(14:0/14:0)", 
#'        "PA(14:0/16:0)", "PA(16:0/0:0)", "PA(16:0/12:0)", "PA(16:0/14:0)",
#'        "PA(16:0/16:0)"),
#'    logFC_cond1 = c(-5.08,  0.75,  5.43, -0.62,  2.35, 1.39, 2.91,  0.26, 
#'        -4.14,  0.19,  6.18, 0.78, -1.81,  4.66, -0.10,  2.84, -0.81,
#'        -0.81, -0.32,  0.17,  2.25, -1.94,  0.80, 4.21,  0.20, -3.29, 
#'        -0.11),
#'    logFC_cond2 = c(-2.73,  6.14,  1.98,  0.09,  1.57,  1.77,  3.08,  4.04,
#'        -3.01, 1.22, -4.25, 0.39, 0.53, 3.30, 7.10, 2.81, -0.99, -0.09,
#'        -8.25, 4.94, -3.54, -7.74, -1.98, 0.73,  2.36,  2.53, -0.62))
#' 
#' ## apply the function
#' add_vertex_attributes(g = g, attributes = attributes_df, col_vertex = "name")
add_vertex_attributes <- function(g, attributes, col_vertex = colnames(attributes)[1]) {
    
    ## check arguments
    if (!is.data.frame(attributes))
        stop("'attributes' has to be a data.frame.")
    
    if (length(col_vertex) != 1)
        stop("'col_vertex' has to be of length 1.")
    
    
    if (!col_vertex %in% colnames(attributes))
        stop("'col_vertex' not in 'colnames(attributes)'.")
    
    if (any(duplicated(attributes[[col_vertex]])))
        stop("'attributes[[col_vertex]]' contains duplicated entries.")
    
    ## subset attributes such that it does not contain entries that are not
    ## vertices of g
    attributes <- attributes[attributes[[col_vertex]] %in% names(igraph::V(g)), ]
    
    ## obtain the colnames, obtain the columns that store attribute information
    cols_df <- colnames(attributes)
    cols_df_attr <- cols_df[cols_df != col_vertex]
    
    ## iterate trough the attribute columns and add attribute information
    for (attr_i in cols_df_attr) {
        ## set V(g) for cols_df_attr to NA
        g <- igraph::set_vertex_attr(graph = g, name = attr_i, value = NA)
        
        ## set V(g) for cols_df_attr to attribute values
        g <- igraph::set_vertex_attr(graph = g, name = attr_i, 
                                     index = attributes[[col_vertex]], 
                                     value = attributes[[attr_i]])   
    }
    
    ## return the graph
    g
    
}
