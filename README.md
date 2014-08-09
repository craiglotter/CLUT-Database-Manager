CLUT-Database-Manager
=====================

Commerce IT currently runs a stat keeping exercise in their labs that logs all login data to a database housed on their webserver. CLUT Database Manager is an application that runs on said server and is responsible for catching all month changes as that is when it needs to duplicate the existing Lab_Usage_Tracker.mdb database and archive its contents for the month that has just passed. Included in this operation is to remove all last month's entries out of this month's database. The reason for this is that the original CLUT (Commerce Lab Usage Tracker) system was rapidly prototyped using a standard MS Access database, meaning that its rather severe limitations quickly becomes a liability.  Created by Craig Lotter, November 2008
