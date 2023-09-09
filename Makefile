
.PHONY: clean

clean:
	rm -rf ./MergeExcels
	rm -rf ./main
	rm -rf ./merged.xlsx
	rm -rf ./*.exe


build:
	go mod tidy
	go build

run:
	go run main.go
