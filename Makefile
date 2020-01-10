

build:
	@echo "Build binaries"
	@GOOS=linux GOARCH=amd64 go build -o ./bin/pentacon_museum-linux-amd64 github.com/kergeodeta/pentaconmuseum
	@GOOS=windows GOARCH=amd64 go build -o ./bin/pentacon_museum-windows-amd64.exe github.com/kergeodeta/pentaconmuseum
