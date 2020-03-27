## pdf2pptx
Go library to convert pdf to pptx via images

### Install

    go get github.com/eletrolitico/pdf2pptx
    
    
### Example
```go
package main

import pptx "github.com/eletrolitico/pdf2pptx"

func main(){
  pptx.ConvertToPptx("MyDoc.pdf", "MyDoc.pptx")
}
```
