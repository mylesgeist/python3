f = open("./amzn1.tortuga.txt", "r")
file_number = 1
CHUNK_SIZE = 1000000
chunk = f.readline()
output = ""

while chunk:
	output += chunk
	chunk = f.readline()
	
	if len(output) > CHUNK_SIZE:
		w = open("AmazonFeed_Errors_split_" + str(file_number) + ".txt", "w")
		w.write(output)
		file_number += 1
		output = ""
		
w = open("AmazonFeed_Errors_split_" + str(file_number) + ".txt", "w")
w.write(output)		

f.close
w.close