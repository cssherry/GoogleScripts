# Google Sheet Layout
The sheet has the following format:

## Form Responses 1 (tab1)
* Timestamp	(column from form)
* Name	(column from form)
* Did you finish the expected book?	(column from form)
* Rate the book you read!	(column from form)
* Which book did you actually finish reading?	(column from form)
* WhoWillReadNext	(column from script)
* WaitingForNewBook?	(column from script)

## Schedule (tab2)
* NewCycle (column)
  * 5/1/2015
  * 7/1/2015
  * 9/1/2015
  * 11/1/2015
  * 1/1/2016
  * 3/1/2016
  * 5/1/2016
  * 7/1/2016
* List of participants (column)

## Addresses (tab3)
* Name	(column from participant)
* Address	(column from participant)
* Email	(column from participant)
* Book (column from participant)
* Choices (column from participant)

# Problem
* This can only be optimized based on reader timing or perfect cycle matching.
* If based on reader timing, there is a chance that the last person to finish reading will not be matched
* If based on perfect cycle matching, there is a chance that someone who reads quickly will have to wait a long time for their assigned match to mail them their book. Also, there is a problem on how to optimize the randomized assignment to create perfect matches.
* Solution, optimize on time and hope for the best. 