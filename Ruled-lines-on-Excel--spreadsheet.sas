/******Excelに罫線を引く******/
%macro dde_cell_format(file=, sheet=, cell=, type=);
%put --------------------------------------------------;
%put dde_cell_format;       /*セルの書式設定をする*/
%put &=file;                /*対象ファイルを指定（省略した場合はアクティブファイルが対象となる）*/
%put &=sheet;               /*対象シートを指定（省略した場合はアクティブシートが対象となる）*/
%put &=cell;                /*結合するセルを指定（R1C1形式のみ指定可、例r3c3:r4c4）*/
%put &=type;                /*書式設定を指定する*/
%put --------------------------------------------------;
       filename xlssys dde "excel|system";
       data _null_;
              file xlssys;
              put '[error(false)]';
              %if %length(%superq(file))>0 %then %do;
                     put "[activate(""%superq(file)"")]";
              %end;
              %if %length(%superq(sheet))>0 %then %do;
                     put "[workbook.activate(""%superq(sheet)"")]";
              %end;
              put "[select(""&cell."")]";
              put "[%superq(type)]";
       run;
       filename xlssys clear;
%mend dde_cell_format;

 

%macro set_go_border(inds=,var1=,var2=,sheet=, offset=,StCol1=,EndCol1=,StCol2=,EndCol2=);
	data _NULL_;
              set &inds. end=_EOF;
                     if  _EOF then call symputx("OBS", _N_);
                     if lag(&var1.) ^= &var1. then do;
                            call symputx(cats("A",_N_), _N_);
                     end;
                     else do;
                            call symputx(cats("A",_N_), "0");
                     end;
                     if lag(&var2.) ^= &var2. then do;
                            call symputx(cats("B",_N_), _N_);
                     end;
                     else  do;
                            call symputx(cats("B",_N_), "0");
                     end;
       run;
/*罫線を引く*/
/*
border(<外枠>,<左>,<右>,<上>,<下>,<網掛け>,<外枠色>,<左色>,<右色>,<上色>,<下色>)
<外枠>,<左>,<右>,<上>,<下>
0:なし1:実線2:太線3:破線4:点線5:極太線6:二重線7:細線
<網掛け>
true or false
<外枠色>,<左色>,<右色>,<上色>,<下色>
1～56で指定する
*/
    %do i=1 %to &OBS;
         %if &&A&i ^=0 %then %do;
                     %dde_cell_format(file=&out. , sheet=&sheet., cell=R%EVAL(&&A&i + &offset.)C&StCol1.:r%EVAL(&&A&i+&offset.)C25, type=border(0,0,0,1,0,false,0,0,0,0,0,0));
              %end;
              %if &&B&i ^=0 %then %do;
                     %dde_cell_format(file=&out. , sheet=&sheet. ,cell=R%EVAL(&&B&i+&offset.)C&StCol2:r%EVAL(&&B&i+&offset.)C25, type=border(0,0,0,1,0,false,0,0,0,0,0,0));
              %end;
       %end;
       /*最後に罫線で閉じる*/
       %dde_cell_format(file=&out. , sheet=&sheet., cell=R%EVAL(&OBS.+&offset.+1)C&StCol1.:R%EVAL(&OBS.+&offset.+1)C25, type=border(0,0,0,1,0,false,0,0,0,0,0,0));

%mend set_go_border;