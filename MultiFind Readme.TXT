This is all free, no warranty, all risk is yours, etc etc blah blah...

This is a simple project, it simply lets you pass a string and a substring in, and it gives you back a count of any matches.

So you pass a target string into the function, and a comma-delimited substring of other string items, and it will return an integer that shows how many of your substring list items were found in the target string.

Call it like this:

    ReturnValue = MultiFind (YourTargetString, YourListOfPossibleMatches)

Here's an example:

    Dim X as Integer
    X = MultiFind ("This here would be your target string", "this,target,string,item,hippopotamus")

    When you execute the code, X receives the value 3 because "this", "target" and "string" were found within the first string.

You can use it as a DLL so it's available to all your programs, even MS Office apps, or just cut and paste the code from the DLL's source and use/modify/etc directly in your VB programs.

Have fun!

Brian Battles WS1O
Middletown, CT USA
brianb@cmtelephone.com