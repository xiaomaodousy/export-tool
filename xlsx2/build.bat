go build -o ../../../scripts/tools/xlsx2csv-win.exe main.go
SET  CGO_ENABLED=0
SET GOOS=darwin
SET GOARCH=amd64
go build -o ../../../scripts/tools/xlsx2csv-mac main.go
SET CGO_ENABLED=0
SET GOOS=linux
SET GOARCH=amd64
go build -o ../../../scripts/tools/xlsx2csv-linux main.go