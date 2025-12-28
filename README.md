# Generate Tithe Receipts

I needed to generate tithes & offerings receipts for church. Also I needed to 
learn Go.

## Usage

First you will need to download a zipfile of tithe & offering giving sheets 
from the Google drive.  You can do this by selecting the year's giving sheet 
folder in the Google Drive and choosing Download.

Once you have the zipfile of giving sheets locally, you can then pass that 
zipfile path as the first argument to the utility.

The second argument you will need to pass to the utility is the docx template
file.

Linux Example:

```
./generate-tithe-receipts ./2025\ Giving\ Sheets-20251226T163304Z-3-001.zip ./donation-receipt-template.docx
```

## Build instructions

### Prerequisities

To build the binaries, you will need a working Go development environment.

You will also need to install a couple Go packages into your environment.

This worked for me:

```
go get github.com/xuri/excelize/v2
go get github.com/lukasjarosch/go-docx
```

### Build Linux binary

To build a Linux binary:

```
go build generate-tithe-receipts.go
```

### Build Windows binary

To build a Windows binary (on Linux):

```
GOOS=windows GOARCH=amd64 go build generate-tithe-receipts.go
```

