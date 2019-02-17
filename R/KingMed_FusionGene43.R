#' To clear 43 fusion genes data from King Med test to a dataframe
#'
#' @param filename the name of fusion gene data text
#' @param WriteExecl A logical object. TRUE as default to write excel
#'
#' @return a dataframe
#' @export
#'
#' @examples KingMed_FusionGene43("data.txt")
KingMed_FusionGene43<-function(filename,WriteExecl=TRUE){
    library(textreadr)
    library(stringr)
    library(stringi)
    library(magrittr)
    library(plyr)
    library(tmcn)
    #read the long string
    if (grepl(".txt",filename)) { TRdata=suppressWarnings(toString(readLines(filename)))
    }else if(grepl(".doc",filename) | grepl(".docx",filename)) {TRdata=toString(read_docx(filename))}
    #standardize characters
    TRdata=stri_trans_nfkd(TRdata)
    TRdata=TRdata %>%
      gsub(pattern = toUTF8("\u533B \u9662"),replacement = toUTF8("\u533B\u9662"))%>%
      gsub(pattern = toUTF8("\u79D1 \u5BA4"),replacement = toUTF8("\u79D1\u5BA4"))%>%
      gsub(pattern = toUTF8("\u6027 \u522B"),replacement = toUTF8("\u6027\u522B"))%>%
      gsub(pattern = toUTF8("\u5E74 \u9F84"),replacement = toUTF8("\u5E74\u9F84"))%>%
      gsub(pattern = toUTF8("CD45\u5F31, \u8868\u8FBE\u7EC6\u80DE"),replacement = toUTF8("CD45\u5F31\u8868\u8FBE\u7EC6\u80DE"))%>%
      gsub(pattern = toUTF8("CD45\u9634\u6027, \u8868\u8FBE\u7EC6\u80DE"),replacement = toUTF8("CD45\u9634\u6027\u8868\u8FBE\u7EC6\u80DE"))%>%
      gsub(pattern = toUTF8("\u5355\u6838\u7EC6, \u80DE"),replacement = toUTF8("\u5355\u6838\u7EC6\u80DE"))%>%
      gsub(pattern = toUTF8("\u7EFC, \u5408\u8003\u8651"),replacement = toUTF8("\u7EFC\u5408\u8003\u8651"))%>%
      gsub(pattern = toUTF8("\u5F62, \u6001\u5B66"),replacement = toUTF8("\u5F62\u6001\u5B66"))%>%
      gsub(pattern = toUTF8("\u5176, \u514D\u75AB\u8868\u578B"),replacement = toUTF8("\u5176\u514D\u75AB\u8868\u578B"))%>%
      gsub(pattern = toUTF8("\u8F85, \u52A9\u8BCA\u65AD\u65B9\u6CD5"),replacement = toUTF8("\u8F85\u52A9\u8BCA\u65AD\u65B9\u6CD5"))%>%
      gsub(pattern = toUTF8("\u5176\u5B83, \u4E34\u5E8A\u4FE1\u606F"),replacement = toUTF8("\u5176\u5B83\u4E34\u5E8A\u4FE1\u606F"))%>%
      gsub(pattern = toUTF8("\u59D3 \u540D"),replacement = toUTF8("\u59D3\u540D"))

    #split
    tongren=toUTF8('\u4E0A\u6D77\u4EA4\u901A\u5927\u5B66\u533B\u5B66\u9662\u9644\u5C5E\u540C\u4EC1\u533B\u9662')
    TRdata1=strsplit(x = TRdata,split = tongren)[[1]]
    #get fusion gene data
    dataget=c()
    fusiongeneTitle=toUTF8("\u767D\u8840\u75C5\u4E2D43\u878D\u5408\u57FA\u56E0\u7B5B\u67E5")
    for (data.i in 1:length(TRdata1)) {
      if (suppressWarnings(grepl(fusiongeneTitle,TRdata1[data.i]))){
        dataget=c(dataget,data.i)
      }
    }
    #
    TR.datafrme=data.frame()
    prgbar<- txtProgressBar(min = 0, max = length(dataget),
                            style = 3,
                            initial = 0,width = 35)
    for (flui.i in 1:length(dataget)) {
      #first page
      TRdata2.i=TRdata1[dataget[flui.i]]
      #Title for two pages
      #sample code
      SampleCode=toUTF8("\u6807\u672C\u6761\u7801")
      SampleCode.Num=gsub(pattern = " ",replacement = "",
                          x=gsub(pattern = SampleCode,replacement = "",
                         x =str_extract(string = TRdata2.i,
                                 pattern = paste0(SampleCode,"[0-9a-zA-Z ]{1,}"))))
      #PatientName
      PatientName=toUTF8("\u75C5\u4EBA\u59D3\u540D")
      PatientName.Num=gsub(pattern = " ",replacement = "",
                          x=sub(pattern = toUTF8("\u79D1\u5BA4"),replacement = "",
                          sub(pattern = toUTF8("\u59D3\u540D"),replacement = "",
                              x =str_extract(string = TRdata2.i,
                                             pattern = paste0(toUTF8("\u59D3\u540D"),".*",toUTF8("\u79D1\u5BA4"))))))
      #Department
      Department=toUTF8("\u79D1\u5BA4")
      Department.Num=gsub(pattern = " ",replacement = "",
                      x=sub(pattern = toUTF8("\u5B9E\u9A8C\u53F7"),replacement = "",
                         sub(pattern = Department,replacement = "",
                             x =str_extract(string = TRdata2.i,
                                            pattern = paste0(Department,".*",toUTF8("\u5B9E\u9A8C\u53F7"))))))
      #testid
      testid=toUTF8("\u5B9E\u9A8C\u53F7")
      testid.Num=gsub(pattern = " ",replacement = "",
                x=sub(pattern = "",replacement = "",
                     sub(pattern = testid,replacement = "",
                         x =str_extract(string = TRdata2.i,
                                        pattern = paste0(testid,"[a-zA-Z0-9 ]*")))))
      #sex
      sex=toUTF8("\u6027\u522B")
      sex.Num=gsub(pattern = " ",replacement = "",
                x=sub(pattern = toUTF8("\u623F/\u5E8A\u53F7"),replacement = "",
                  sub(pattern = sex,replacement = "",
                      x =str_extract(string = TRdata2.i,
                                     pattern = paste0(sex,".*",toUTF8("\u623F/\u5E8A\u53F7"))))))
      #bedid
      bedid=toUTF8("\u623F/\u5E8A\u53F7")
      bedid.Num=gsub(pattern = " ",replacement = "",
                  x=sub(pattern = toUTF8("\u95E8\u8BCA/\u4F4F\u9662\u53F7"),replacement = "",
                    sub(pattern = bedid,replacement = "",
                        x =str_extract(string = TRdata2.i,
                                       pattern = paste0(bedid,".*",toUTF8("\u95E8\u8BCA/\u4F4F\u9662\u53F7"))))))
      #Hospitalid
      Hospitalid=toUTF8("\u95E8\u8BCA/\u4F4F\u9662\u53F7")
      Hospitalid.Num=gsub(pattern = " ",replacement = "",
                      x=sub(pattern =  paste0('[ ,]*',toUTF8("\u5E74\u9F84")),replacement = "",
                         sub(pattern = Hospitalid,replacement = "",
                             x =str_extract(string = TRdata2.i,
                                            pattern = paste0(Hospitalid,".*", toUTF8("\u5E74\u9F84"))))))
      #age
      age=toUTF8("\u5E74\u9F84")
      age.Num=gsub(pattern = " ",replacement = "",
              x=sub(pattern =  paste0('[ ,]*',toUTF8("\u7533\u8BF7\u533B\u751F")),replacement = "",
                  sub(pattern = age,replacement = "",
                      x =str_extract(string = TRdata2.i,
                                     pattern = paste0(age,".*", toUTF8("\u7533\u8BF7\u533B\u751F"))))))
      #AcceptTime
      AcceptTime=toUTF8('\u63A5\u6536\u65F6\u95F4')
      AcceptTime.Num=sub(pattern = " ",replacement = "",
                    x=sub(pattern =  paste0('[ ,]*',toUTF8("\u9001\u68C0\u6807\u672C")),replacement = "",
                         sub(pattern = toUTF8("\u63A5\u6536\u65E5\u671F"),replacement = "",
                             x =str_extract(string = TRdata2.i,
                                            pattern = paste0(toUTF8("\u63A5\u6536\u65E5\u671F"),".*", toUTF8("\u9001\u68C0\u6807\u672C"))))))
      #ApplyDoctor
      ApplyDoctor=toUTF8("\u7533\u8BF7\u533B\u751F")
      ApplyDoctor.Num=sub(pattern = " ",replacement = "",
                      x=sub(pattern =  paste0('[ ,]*',toUTF8("\u63A5\u6536\u65E5\u671F")),replacement = "",
                          sub(pattern = ApplyDoctor,replacement = "",
                              x =str_extract(string = TRdata2.i,
                                             pattern = paste0(ApplyDoctor,".*", toUTF8("\u63A5\u6536\u65E5\u671F"))))))
      #bacterial
      bacterial=toUTF8("\u9001\u68C0\u6750\u6599")
      bacterial.Num=sub(pattern = " ",replacement = "",
                    x=sub(pattern =  paste0('[ ,]*',toUTF8("\u533B\u9662\u6807\u8BC6")),replacement = "",
                        sub(pattern = toUTF8("\u9001\u68C0\u6807\u672C"),replacement = "",
                            x =str_extract(string = TRdata2.i,
                                           pattern = paste0(toUTF8("\u9001\u68C0\u6807\u672C"),".*", toUTF8("\u533B\u9662\u6807\u8BC6"))))))
      #SamplingTime
      SamplingTime=toUTF8("\u91C7\u6837\u65F6\u95F4")
      SamplingTime.Num=gsub(pattern =  '[a-zA-Z ,]*',replacement = "",
                            sub(pattern = toUTF8("\u91C7\u6837\u65E5\u671F"),replacement = "",
                                x =str_extract(string = TRdata2.i,
                                               pattern = paste0(toUTF8("\u91C7\u6837\u65E5\u671F"),"[a-zA-Z0-9- ]*"))))
      #SamplingCondition
      SamplingCondition=toUTF8("\u6807\u672C\u60C5\u51B5")
      SamplingCondition.Num=gsub(pattern = " ",replacement = "",
                          x=sub(pattern =  toUTF8("\u8054\u7CFB\u7535\u8BDD"),replacement = "",
                                sub(pattern = SamplingCondition,replacement = "",
                                    x =str_extract(string = TRdata2.i,
                                                   pattern = paste0(SamplingCondition,".*",toUTF8("\u8054\u7CFB\u7535\u8BDD"))))))
      wholeresult=toUTF8("\u68C0\u6D4B\u7ED3\u679C")
      wholeresultlocation=gregexpr(wholeresult,TRdata2.i)[[1]][1]
      firstfglocation=gregexpr(toUTF8("\u878D\u5408\u57FA\u56E0:"),TRdata2.i)[[1]][1]
      wholeresult.Num=substring(TRdata2.i,wholeresultlocation,firstfglocation-1)%>%
        sub(pattern = paste0(" .*",toUTF8("\u6cd5")," "),replacement = ";")%>%
        gsub(pattern = " ",replacement = "")
      #fusion genes in first page
      fusiongeneframe=c()
      fusiongenepage2=c()
      fusiongeneframe=TRdata2.i %>%
        substring(first = firstfglocation,last = nchar(TRdata2.i))%>%
        gsub(pattern = paste0(" .{,4}",toUTF8("\u6cd5")),replacement = "")%>%
        gsub(pattern = toUTF8("\u878D\u5408\u57FA\u56E0:"),replacement = "")%>%
        gsub(pattern = paste0(")[, ]*",toUTF8("\u672C\u68C0\u6D4B"),".*"),replacement = "),")%>%
        gsub(pattern = ", ",replacement = ",")%>%
        gsub(pattern = " ",replacement = ";")

      fusiongenepage2=TRdata1[dataget[flui.i]+1]%>%
        sub(pattern = paste0(".*",toUTF8("\u53C2\u8003\u503C,")),replacement = "")%>%
        sub(pattern = paste0("[ ,]{0,}",toUTF8("\u5EFA\u8BAE\u4E0E\u89E3\u91CA:"),".*"),replacement = "")%>%
        gsub(pattern = paste0(" .{,4}",toUTF8("\u6cd5")),replacement = "")%>%
        gsub(pattern = toUTF8("\u878D\u5408\u57FA\u56E0:"),replacement = "")%>%
        gsub(pattern = paste0(")[, ]*",toUTF8("\u672C\u68C0\u6D4B"),".*"),replacement = "),")%>%
        sub(pattern = " {0,}",replacement = "")%>%
        gsub(pattern = ", ",replacement = ",")%>%
        gsub(pattern = " ",replacement = ";")
      allfusiongene=paste0(wholeresult.Num,fusiongeneframe,fusiongenepage2)
      fusiongenelist=strsplit(x = allfusiongene,split = ",")
      fusiongenelist2=strsplit(x = fusiongenelist[[1]],split = ";")
      fgresult=c()
      for (fgunlist.i in fusiongenelist2) {
        fgunlist.clear= fgunlist.i[2] %>%
          sub(pattern = ".*\\(",replacement = "")%>%
          sub(pattern = "\\)",replacement = "")
        fgresult=c(fgresult,fgunlist.clear)
      }
      fgresult=ifelse(fgresult=="+",toUTF8("\u9633\u6027"),fgresult)
      fgtitle=c()
      for (fgunlist.i in fusiongenelist2) {
        fgtitle=c(fgtitle,fgunlist.i[1])
      }
      fgfinaledata=data.frame(matrix(data = fgresult,nrow = 1))
      colnames(fgfinaledata)=fgtitle

      ##########
      #dataframe
      TR.datafrme.i=cbind(SampleCode.Num,
                          PatientName.Num,
                          Department.Num,
                          testid.Num,
                          sex.Num,
                          bedid.Num,
                          Hospitalid.Num,
                          age.Num,
                          AcceptTime.Num,
                          ApplyDoctor.Num,
                          bacterial.Num,
                          SamplingTime.Num,
                          SamplingCondition.Num)
      colnames(TR.datafrme.i)= c(SampleCode,
                               PatientName,
                               Department,
                               testid,
                               sex,
                               bedid,
                               Hospitalid,
                               age,
                               AcceptTime,
                               ApplyDoctor,
                               bacterial,
                               SamplingTime,
                               SamplingCondition)
      TR.datafrme.bind=cbind(TR.datafrme.i,fgfinaledata)
      TR.datafrme=rbind.fill(TR.datafrme,TR.datafrme.bind)
      setTxtProgressBar(pb = prgbar, value = flui.i)
    }
    close(prgbar)
    TR.datafrme1=data.frame(ifelse(is.na(as.matrix(TR.datafrme)),"",as.matrix(TR.datafrme)))
    colnames(TR.datafrme1)=colnames(TR.datafrme)
    TR.datafrme=TR.datafrme1
    cat( toUTF8("\u68C0\u6D4B\u9879\u76EE:"),
         "\n",
         toUTF8("43\u79CD\u878D\u5408\u57FA\u56E0")
         )
    if (WriteExecl==TRUE){
      filenamewritten=Sys.time() %>%
        sub(pattern = ":",replacement = toUTF8("\u02B1"))%>%
        sub(pattern = ":",replacement = toUTF8("\u5206"))%>%
        paste0(toUTF8("\u79D2 43\u79CD\u878D\u5408\u57FA\u56E0.csv"))
      write.csv(TR.datafrme,filenamewritten)
      cat("\n","\n",
          toUTF8("\u6570\u636E\u88AB\u5199\u5165Excel\u4E2D,\u6587\u4EF6\u540D\u4E3A:"),
          "\n",
          filenamewritten)
    }
    return(TR.datafrme)
}
