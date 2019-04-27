using namespace DocumentFormat.OpenXml.Wordprocessing
using namespace DocumentFormat.OpenXml.Spreadsheet
using namespace DocumentFormat.OpenXml.Packaging
using namespace DocumentFormat.OpenXml
function OpenXMLItem{
    Param([string]$type,
          [Scriptblock]$contents={},
          $object)
    if($type){
      $t=New-Object $type
      write-verbose "Constructing a $type"
    } else {
      $t=$object
    }
    $obj=$null
    [array]$obj=(& $contents)
    if($obj -ne $null ){
    write-verbose "found $($obj.Count) items to add to $type"
 
    foreach($item in $obj){
       if($item){
            if($item.Type -eq 'Row'){
              write-verbose "row"
            }
            write-verbose "Adding item to $type of type $($item.o.GetType().Name)"
         [void]$t.AppendChild($item.o);
       }
    }
    }
    write-verbose "Returning an object of type $($t.GetType().Name)"
    write-output $t -NoEnumerate
 
}