Attribute VB_Name = "BlockReembeddingList"
'holds the data of the blocks for reembedding

'current start index of the block group presented on screen
Public iGuiGroup As Integer

'max number of lines to load on screen
Public Const MAX_LINES = 6

'the list of BlockReembedding items
Public d As New Dictionary
