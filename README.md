# Printer Audit and Cleanup Utility

Batch script for checking and cleaning local printers, network printers, print server connections, ports, and printer drivers on a Windows client.

## Functions

- **List installed printers**  
  Shows all locally installed and connected printers.

- **List printer ports**  
  Displays TCP/IP and other printer ports.

- **List printer drivers**  
  Shows installed printer drivers and driver details.

- **Search printer / port / driver**  
  Searches for a custom term such as a printer name, port name, driver name, or comment.

- **Check print server connections**  
  Displays printers that appear to come from a print server or printer connection deployment.

- **Check registry connections**  
  Reviews user-based and computer-based printer connection registry entries.

- **Check spooler status**  
  Shows the current print spooler service state.

- **Generate full report**  
  Creates a TXT report in:
  `Desktop\PrinterReports`

- **Delete printer by exact name** *(Admin)*  
  Removes a specific printer queue.

- **Delete printer port by exact name** *(Admin)*  
  Removes a specific printer port.

- **Delete network printer connection by UNC**  
  Removes a user-based printer connection such as `\\PRINTSERVER\PrinterName`.

- **Delete computer printer connection by UNC** *(Admin)*  
  Removes a computer-based printer connection.

- **Restart spooler** *(Admin)*  
  Restarts the print spooler service after cleanup.

## Notes

- If printers are deployed centrally by **print server** or **GPO**, local cleanup alone may not be enough.
- Stale printers should usually be removed in this order:
  1. Printer queue
  2. Printer port
 

  4. Driver only if still necessary
- Administrator rights are required for queue, port, and spooler changes.
