Attribute VB_Name = "BlockAddSlideList"
    'holds the data of the blocks to add slides from

'current start index of the block group presented on screen
Public iGuiGroup As Integer

'max number of lines to load on screen
Public Const MAX_LINES = 4

'the list of BlockAddSlide items
Public d As New Dictionary
