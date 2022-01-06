# edit-dynamically-doc-file-using-apache-poi

# the Flow:
#  the CCQMContractTemplate.docx has a lot of blanks for infomations. a unique variable is put on each blank (we set it in camel case)
#  we are going to create a map with its keys named after the variables in the docx and its values are what will be filled in the blank
  After programme runs, there are 3 files generated in folder template (write some codes to auto-delete them if you want to or do it manually):
          1, contract.docx is our contract after filled by the values in the map
          2, base64.txt is our text file which contains base64-code encoded from contract.docx
          3, encoded-contract.docx is the doc file decoded from base64-code comes from base64.txt file
          4, actually there is one more file as a clone file and actually the contract.docx are going to be generated from this clone
             not from the template. Because our method would overwrite the values into the template so a clone have to be created for
             sort of sacrificing. It will be created then deleted in the blink of an eye that we can not see its existence in the project's tree.
             I know there are many samples that can manage the project without generating a clone file but I find it difficult
             and complicated for newbies. Everything will be done with this short-and-sweet method with a tiny clone file.

  The values we'll use here are static not from a database. Create one and do custom your own values in value-part of the map below as your requests
 
