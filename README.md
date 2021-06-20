# Celebrate
PowerPoint macro for generating work anniversary and birthday slides.

Celebrate is compatible with Workday®'s RaaS json output. See information on creating the json with Workday RaaS below.

![image](https://user-images.githubusercontent.com/413552/122630517-309abe80-d079-11eb-8882-364dc13029ee.png)

## Installation
* View the latest release - https://github.com/swhitley/Celebrate/releases/latest
* Download `Celebrate.pptm` and `data.txt` to the same directory.  Open `Celebrate.pptm` and follow the instructions.


## Sample Data

This is an example of the json format that Celebrate can process.  Celebrate also recognizes the json that is produced by Workday's Reporting as a Service (RaaS).

```json
[
	{
		"group": "5",
		"photo": "https://tile.loc.gov/image-services/iiif/service:music:musgottlieb:musgottlieb-00151:ver01:0001/full/pct:25.0/0/default.jpg",
		"name": "Louis Armstrong",
		"title": "Satchmo'"
	},
	{
		"group": "5",
		"photo": "https://tile.loc.gov/image-services/iiif/service:music:musgottlieb:musgottlieb-04291:ver01:0001/full/pct:25.0/0/default.jpg",
		"name": "Lena Horne",
		"title": "Jazz Woman"
	}
]
```

## About the Json Data
* `photo` can be a url, file path, or a base64-encoded image.  Use double-backslashes (\\\\) for local files on Windows (e.g. c:\\\\temp\\\\pic.jpg).  
  
  Note: If you are viewing the text version of this README file, you will see four backslashes instead of two. This is due to markdown requirements for this file. Only use two backslashes in the data.

## Control Slide
Run the macro from this slide.
* Change the `values` in the `options` table to change titles, labels, and colors.
* Enter presentation mode.
* Click the document image to open the `data.txt` file.
* Paste the Json into `data.txt`. Save and close the file.
* Click the `Run` button.

![image](https://user-images.githubusercontent.com/413552/122657872-0f3ede80-d11c-11eb-85be-d302dc764581.png)



## Workday RaaS Configuration

Create a Workday report to generate the json data for Celebrate.

![image](https://user-images.githubusercontent.com/413552/122632795-e7069f80-d089-11eb-86a7-005757839e99.png)

![image](https://user-images.githubusercontent.com/413552/122632897-7ca22f00-d08a-11eb-9413-b15471d13140.png)

![image](https://user-images.githubusercontent.com/413552/122632963-cf7be680-d08a-11eb-88ca-0be1db14e83d.png)

![image](https://user-images.githubusercontent.com/413552/122633079-80828100-d08b-11eb-923a-753a884192db.png)


## Credits
WebHelpers
(c) Tim Hall - https://github.com/VBA-tools/VBA-Web (MIT License)

Celebrate is compatible with Workday®
It is not sponsored, affiliated with, or endorsed by Workday.

