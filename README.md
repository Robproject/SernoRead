Logic:
```
1. get file_dir  
2. get xlsx  
3. get sernos  
	a. for file in dir  
		i. if name has 'input'  
			I. find serno text  
				1. get pcb image area / threshold  
				2. compare width/height  
					a. rotate 90 if horizontal  
				3. check bottom left for bt1 placement  
					a. rotate 180 if present  
				4. crop top edge of board with text  
				5. ocr  
					a. expand image size  
					b. tesseract read  
			II. create dict (filename : serno)  
4. write sernos to xlsx  
	I. for key in serno dict  
		a. if filename (key) in xlsx dict, compare values:  
			i. if same, skip and notify  
			ii. if value empty, write  
			iii. if value different, skip and notify  
		b. else, create and write  


```
