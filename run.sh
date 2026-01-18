#!/bin/bash

# Navigate to the source directory
cd src

# Ensure dependencies are up to date
go get -u github.com/chromedp/chromedp
go mod tidy

# Run the converter
go run .