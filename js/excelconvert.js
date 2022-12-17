
  var foir=0;
  var disposable=0;
  const excel_file = document.getElementById('excel_file');
  var trial_name;
  excel_file.addEventListener('change', (event) => {
  
      if(!['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'].includes(event.target.files[0].type))
      {
          document.getElementById('excel_data').innerHTML = '<div class="alert alert-danger">Only .xlsx or .xls file format are allowed</div>';
  
          excel_file.value = '';
  
          return false;
      }

      var reader = new FileReader();
  
      reader.readAsArrayBuffer(event.target.files[0]);
  
      reader.onload = function(event){
        for(let i=0; i<=2; i++){
          var data = new Uint8Array(reader.result);
  
          var work_book = XLSX.read(data, {type:'array'});
  
          var sheet_name = work_book.SheetNames;
  
          var sheet_data = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[i]], {header:1});
          
          console.log(sheet_data);

          if(sheet_data.length > 0){
            
            if(i==0){
              
              var count=0;
              var table_output = '<table class="table table-striped table-bordered">';
                for(var row = 0; row < sheet_data.length-1; row++){
                  table_output += '<tr>';
                    for(var cell = 0; cell < sheet_data[row].length; cell++)
                    {
                        
                        if(row == 0)
                        {
    
                          table_output += '<th>'+sheet_data[row][cell]+'</th>';
    
                        }
                        
                        else if(row>20){
                          
                          if(count==0){
                            
                            if(row==22)
                            {
                              table_output+='</table><div class="html2pdf__page-break"  style="margin-top: 20%; margin-bottom:20%;" ></div>';
                              console.log("Inside R22");
                              for(var cell = 0; cell < sheet_data[row].length; cell++)
                                  table_output += '<table><th>'+sheet_data[0][cell]+'</th>';
                              count=1;
                              continue;
                            }
                              
                            console.log("Outside R22");
                            
                            for(var cell = 0; cell < sheet_data[row].length; cell++)
                              table_output += '<td>'+sheet_data[21][cell]+'</td>';
                            table_output += '<tr>';
                            for(var cell = 0; cell < sheet_data[row].length; cell++)
                              table_output += '<td>'+sheet_data[22][cell]+'</td>';
                            
                            
                            
                              
                          }
                          else{
                            table_output += '<td>'+sheet_data[row][cell]+'</td>';
                          }
                          
                          console.log(count);
                        }
                        else
                        {
                            table_output += '<td>'+sheet_data[row][cell]+'</td>';
                        }
                    }
                }
                table_output += '</tr>';
                table_output += '</table>';

            let file_path = "excel_data_"
              file_path_final = file_path.concat(i.toString());
              console.log(file_path_final)
              document.getElementById(file_path_final).innerHTML = table_output;
            }

            if(i==1){
              var table_output = '<table class="table table-striped table-bordered">';
                for(var row = 0; row < sheet_data.length; row++){
                  table_output += '<tr>';
                for(var cell = 0; cell < sheet_data[row].length-2; cell++)
                      {
      
                          if(row == 0)
                          {
      
                              table_output += '<th>'+sheet_data[row][cell]+'</th>';
      
                          }
                          else
                          {
      
                              table_output += '<td>'+sheet_data[row][cell]+'</td>';
      
                          }
      
                      }
                      table_output += '</tr>';
            }
            table_output += '</table>';

            let file_path = "excel_data_"
              file_path_final = file_path.concat(i.toString());
              console.log(file_path_final)
              document.getElementById(file_path_final).innerHTML = table_output;

            var row_length=sheet_data.length-1;
              foir=sheet_data[row_length][6];
              disposable=sheet_data[row_length][7];
              existing_emi=sheet_data[row_length][3];
              console.log(foir,disposable,existing_emi);
              
          }
          if(i==2){
            var row_length=sheet_data.length-1;
            name=sheet_data[row_length][0];
            trial_name=name;
            credit_score=sheet_data[row_length][1];
            salary=sheet_data[row_length][2];
            console.log(name,credit_score,salary);

            document.getElementById("name").innerHTML = name;

            document.getElementById("credit_score_element").innerHTML= credit_score;
            document.getElementById("salary_element").innerHTML= salary;
            document.getElementById("foir_amount_element").innerHTML= foir;
            document.getElementById("existing_emi_element").innerHTML= existing_emi;
            document.getElementById("eligible_amount_element").innerHTML= disposable;

            document.getElementById("foirB").innerHTML= foir;
            document.getElementById("emiB").innerHTML= existing_emi;
            document.getElementById("elaB").innerHTML= disposable;
            
            if(disposable<0){
              document.getElementById("elaD").style.color = "red";
              document.getElementById("myP").style.color = "red";
              document.getElementById("para").innerHTML = "Your existing EMI's are more than your calculated FOIR hence your eligible amount is negative"
            }
            if(disposable>0){
              
              document.getElementById("elaD").style.color = "green";
              document.getElementById("myP").style.color = "green";
              document.getElementById("para").innerHTML = "Your existing EMI's are less than your calculated FOIR hence your eligible amount is positive"
            }

          }

          }
        }


  var data = new Uint8Array(reader.result);
  
  var work_book = XLSX.read(data, {type:'array'});

  var sheet_name = work_book.SheetNames;

  


  var sheet_delinquencies = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[0]], {header:1});
  
  console.log("Del =",sheet_delinquencies);
  var row_length=sheet_delinquencies.length;
  var column_length=6;
  console.log("Deli =",column_length);
  
  deli=sheet_delinquencies[row_length-1][column_length];
  console.log(deli);
  if(deli>0){
    console.log("sda"+deli);
    var gfg_down = document.getElementById("case1");
    gfg_down.parentNode.removeChild(gfg_down);
    gfg_down = document.getElementById("case2");
    gfg_down.parentNode.removeChild(gfg_down);
    
    gfg_down = document.getElementById("case4");
    gfg_down.parentNode.removeChild(gfg_down);
    
    gfg_down = document.getElementById("case5");
    gfg_down.parentNode.removeChild(gfg_down);

    gfg_down = document.getElementById("recommendation_text");
    gfg_down.parentNode.removeChild(gfg_down);
    var sheet_data_4 = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[4]], {header:1});
    document.getElementById("list1").innerHTML=sheet_data_4[1][0];
    
  }
  else if(salary==0){
    
    var gfg_down = document.getElementById("case1");
    gfg_down.parentNode.removeChild(gfg_down);
    gfg_down = document.getElementById("case2");
    gfg_down.parentNode.removeChild(gfg_down);
    
    gfg_down = document.getElementById("case3");
    gfg_down.parentNode.removeChild(gfg_down);
    
    gfg_down = document.getElementById("case4");
    gfg_down.parentNode.removeChild(gfg_down);
    gfg_down = document.getElementById("recommendation_text");
    gfg_down.parentNode.removeChild(gfg_down);
  }
  else{
    var sheet_data_3 = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[3]], {header:1});
  console.log(sheet_data)
  document.getElementById("recommendation_text").innerHTML=sheet_data_3[1][0];
   if(sheet_name.length == 6){
    var gfg_down = document.getElementById("case1");
    gfg_down.parentNode.removeChild(gfg_down);
    gfg_down = document.getElementById("case3");
    gfg_down.parentNode.removeChild(gfg_down);
    gfg_down = document.getElementById("case4");
    gfg_down.parentNode.removeChild(gfg_down);

    
    gfg_down = document.getElementById("case5");
    gfg_down.parentNode.removeChild(gfg_down);

    var sheet_data_4 = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[4]], {header:1});
    
    //3,4
    table_output="<table>"
    table_output+=" <tr><th colspan='2'>Case 1</th></tr>"
    for(var row = 1; row < sheet_data_4.length; row++){
                  table_output += '<tr>';
                for(var cell = 0; cell < sheet_data_4[row].length; cell++)
                      {
                              table_output += '<td>'+sheet_data_4[row][cell]+'</td>';
                      }
                      
            }
            table_output += '</tr>';
            table_output+="</table>";
            document.getElementById("disposabel_table3").innerHTML = table_output;
    

    table_output="<table>";
            var sheet_data_5 = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[5]], {header:1});
    
    table_output=" <tr><th colspan='2'>Case 2</th></tr>"
    for(var row = 1; row < sheet_data_5.length; row++){
                  table_output += '<tr>';
                for(var cell = 0; cell < sheet_data_5[row].length; cell++)
                      {
                              table_output += '<td>'+sheet_data_5[row][cell]+'</td>';
                      }
                      
            }
            table_output += '</tr>';
            table_output+="</table>"
            document.getElementById("disposabel_table4").innerHTML = table_output;
            
  }

  else if(sheet_name.length == 5){
    console.log('salaary'+salary);
    //2
    var gfg_down = document.getElementById("case3");
    gfg_down.parentNode.removeChild(gfg_down);
    gfg_down = document.getElementById("case2");
    gfg_down.parentNode.removeChild(gfg_down);
    
    gfg_down = document.getElementById("case4");
    gfg_down.parentNode.removeChild(gfg_down);

    
    gfg_down = document.getElementById("case5");
    gfg_down.parentNode.removeChild(gfg_down);

    table_output="<table>";
            var sheet_data_5 = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[4]], {header:1});
    case1.style.display = "block";
    table_output=" <tr><th colspan='2'>Case 1</th></tr>"
    for(var row = 1; row < sheet_data_5.length; row++){
                  table_output += '<tr>';
                for(var cell = 0; cell < sheet_data_5[row].length; cell++)
                      {
                              table_output += '<td>'+sheet_data_5[row][cell]+'</td>';
                      }
                      
            }
            table_output += '</tr>';
            table_output+="</table>"
            document.getElementById("disposabel_table2").innerHTML = table_output;
  }
  if(salary<20000){
    console.log('gaandi');
    if(sheet_name.length <= 5){
    gfg_down = document.getElementById("case1");
    gfg_down.parentNode.removeChild(gfg_down);
    }
    if(sheet_name.length <= 6){
    gfg_down = document.getElementById("case2");
    gfg_down.parentNode.removeChild(gfg_down);
    }
    if(salary==0){
       
    gfg_down = document.getElementById("case5");
    gfg_down.parentNode.removeChild(gfg_down);
    }
    else{
      
    gfg_down = document.getElementById("case7");
    gfg_down.parentNode.removeChild(gfg_down);
    }
    /*
    if(credit_score<650){
    gfg_down = document.getElementById("case6");
    gfg_down.parentNode.removeChild(gfg_down);
    }
    */  
    gfg_down = document.getElementById("recommendation_text");
    gfg_down.parentNode.removeChild(gfg_down);
  }
  else if(credit_score<650){
    
    if(sheet_name.length == 5){
      gfg_down = document.getElementById("case1");
      gfg_down.parentNode.removeChild(gfg_down);
    }
    if(sheet_name.length == 6){
      gfg_down = document.getElementById("case2");
      gfg_down.parentNode.removeChild(gfg_down);
    }
    
    gfg_down = document.getElementById("recommendation_text");
    gfg_down.parentNode.removeChild(gfg_down);
  }
  else if(disposable<0){
    
    var gfg_down = document.getElementById("case3");
    gfg_down.parentNode.removeChild(gfg_down);
    gfg_down = document.getElementById("case2");
    gfg_down.parentNode.removeChild(gfg_down);
    gfg_down = document.getElementById("case1");
    gfg_down.parentNode.removeChild(gfg_down);
    
    gfg_down = document.getElementById("case5");
    gfg_down.parentNode.removeChild(gfg_down);


  }


  }

      }
  excel_file.style.display= "none";

  });
 
  const date = new Date();

    let day = date.getDate();
    let month = date.getMonth() + 1;
    let year = date.getFullYear();

    let currentDate = `${day}-${month}-${year}`;
    console.log(currentDate);

  $('.btn-download').click(function(e){
    const options = {
      margin: 0,
      filename:trial_name+'_'+ currentDate+'_CAR.pdf',
      image: { 
        type: 'jpeg', 
        quality: 1
      },
      html2canvas: { 
        scale: 1
      },
      jsPDF: { 
        unit: 'in', 
        format: 'a3', 
        orientation: 'portrait' 
      }
    }

    e.preventDefault();
    const element = document.getElementById('invoice');
    html2pdf().from(element).set(options).save();

  });


