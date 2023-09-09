
.PHONY: clean

clean:
	rm -rf ./MergeExcels
	rm -rf ./main
	rm -rf ./merged.xlsx


build:
	go mod tidy
	go build

run:
	go run main.go
