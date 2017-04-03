# SharePoint-Document-Version-Count
This utility will count all files versions for all documents in a SharePoint document library. It also has the possibility to purge all versions using the SPFileVersionCollection.DeleteAll() method.
It is meant to be run at the server in a command prompt.

## Usage  
`CountFileVersions.exe -url <site collection URL> -web <sub site URL> -doclib <name of the document library>`

## Example
`CountFileVersions.exe -url http://intranet -doclib "Shared Documents"`

To output the result to a text file:  
`CountFileVersions.exe -url http://intranet -doclib "Shared Documents" -output result.txt`

The default separation character is the tab char, but it can be replaced with any char with the -sepchar argument:  
`CountFileVersions.exe -url http://intranet -doclib "Shared Documents" -output result.txt -sepchar ;`

To purge all versions:  
`CountFileVersions.exe -url http://intranet -doclib "Shared Documents" -purge`


## Warning
Use this utility with care. If used with the -purge argument, it will permanently delete your documents (old versions) from the library/list.

## Disclaimer
By using this utility you are fully responsible for what it does to your site collection
