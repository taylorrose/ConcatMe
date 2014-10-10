ConcatMe
========

Simple Excel VBA function to concatenate values in a range of cells and add delimiters to the front or back of values  


`ConcatMe`(`Range`,`startDelim`,`endDelim`,`cutString`)

`Range` - Range of cells you wish to concatinate , *e.g.*, `A3:A11`<br>
`startDelim` - String you would like to add to the beginning of each value in the concatination, *e.g.* , `"{"`<br>
`endDelim` - String you woul dlike to add to the end of each value in the concatination, *e.g.* , `"},"`<br>
`cutString` *(Optional)* - If set to `TRUE` then will remove the last character of the `endDelim` on the last value of the series, *e.g.* , `{1111},{2222},{3333},` ---> `{1111},{2222},{3333}`
