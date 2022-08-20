There are a few important notes if you are looking at this project.

This project implements the COM interop for Excel. That is to say, Mac OS will NOT support excel data logging out of the box.
Instead, if you do not need logging I suggest removing Line 6 and any reference to ExcelHandler (class included).
Otherwise, I highly suggest replacing the functions of ExcelHandler with OpenXML.

As to why I used a Windows only implementation, its the one I have the most experience using.

The project also includes a copy of a dll called Arguments. This, as stated in the project itself, is a personal implementation I use frequently in professional projects, so I would appreciate it if it wasn't stolen and used elsewhere without my knowledge.


# Observed Performance
## Results
_These tests were done using the standard settings of the program. 100 Iterations of 1 sequential, 1 parallel, and 1 parallel nested parallel._
![](https://i.imgur.com/oWz3FIC.png)

### Processor
Intel Core i5 - 9600K (Coffee Lake-S)
 - 6 Physical with 6 Threads
 - (At the time of testing) 4100Mhz
 - L1 
	 - Instruction
		 - 6×32 KBytes
		 - TLB : 2MB/4MB, Fully associative, 8 entries
	 - Data
		 - 6×32 KBytes
		 - TLB : 4KB pages, 4-way set associative, 64 entries
 - L2 : 6×256 KBytes
 - L3 : 9MBytes