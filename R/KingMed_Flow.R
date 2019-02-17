#' To clear flow data from King Med test to a dataframe
#'
#' @param filename the name of flow data text
#' @param WriteExecl A logical object. TRUE as default to write excel
#' @return a dataframe
#' @export
#'
#' @examples KingMed_Flow("data.txt")
KingMed_Flow<-function(filename,WriteExecl=TRUE){
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
      gsub(pattern = toUTF8("\u5176\u5B83, \u4E34\u5E8A\u4FE1\u606F"),replacement = toUTF8("\u5176\u5B83\u4E34\u5E8A\u4FE1\u606F"))

    #split
    tongren=toUTF8('\u4E0A\u6D77\u4EA4\u901A\u5927\u5B66\u533B\u5B66\u9662\u9644\u5C5E\u540C\u4EC1\u533B\u9662')
    TRdata1=strsplit(x = TRdata,split = tongren)[[1]]
    #get flow data
    dataget=c()
    flowTitle=toUTF8("\u6D41\u5F0F\u7EC6\u80DE\u514D\u75AB\u8367\u5149\u5206\u6790\u7ED3\u679C\u62A5\u544A\u5355")
    for (data.i in 1:length(TRdata1)) {
      if (suppressWarnings(grepl(flowTitle,TRdata1[data.i]))){
        dataget=c(dataget,data.i)
      }
    }
    TRdata1.2=TRdata1[dataget]
    #get flow data page
    pageget=c()
    for (datapage.i in 1:length(TRdata1.2)) {
      if (grepl("Conclusion",TRdata1.2[datapage.i])){
        pageget=c(pageget,datapage.i)
      }
    }
    TRdata2=TRdata1.2[pageget]
    #
    TR.datafrme=data.frame()
    prgbar<- txtProgressBar(min = 0, max = length(TRdata2),
                            style = 3,
                            initial = 0,width = 35)
    for (flui.i in 1:length(TRdata2)) {
      TRdata2.i=TRdata2[flui.i]
      #sample code
      SampleCode=toUTF8("\u6807\u672C\u6761\u7801")
      SampleCode.Num=sub(pattern = SampleCode,replacement = "",
                         x =str_extract(string = TRdata2.i,
                                 pattern = paste0(SampleCode,"[0-9a-zA-Z ]{1,}")))
      #PatientName
      PatientName=toUTF8("\u75C5\u4EBA\u59D3\u540D")
      PatientName.Num=sub(pattern = toUTF8("\u79D1\u5BA4"),replacement = "",
                          sub(pattern = PatientName,replacement = "",
                              x =str_extract(string = TRdata2.i,
                                             pattern = paste0(PatientName,".*",toUTF8("\u79D1\u5BA4")))))
      #Department
      Department=toUTF8("\u79D1\u5BA4")
      Department.Num=sub(pattern = toUTF8("\u5B9E\u9A8C\u53F7"),replacement = "",
                         sub(pattern = Department,replacement = "",
                             x =str_extract(string = TRdata2.i,
                                            pattern = paste0(Department,".*",toUTF8("\u5B9E\u9A8C\u53F7")))))
      #testid
      testid=toUTF8("\u5B9E\u9A8C\u53F7")
      testid.Num=sub(pattern = "",replacement = "",
                     sub(pattern = testid,replacement = "",
                         x =str_extract(string = TRdata2.i,
                                        pattern = paste0(testid,"[a-zA-Z0-9 ]*"))))
      #sex
      sex=toUTF8("\u6027\u522B")
      sex.Num=sub(pattern = toUTF8("\u623F/\u5E8A\u53F7"),replacement = "",
                  sub(pattern = sex,replacement = "",
                      x =str_extract(string = TRdata2.i,
                                     pattern = paste0(sex,".*",toUTF8("\u623F/\u5E8A\u53F7")))))
      #bedid
      bedid=toUTF8("\u623F/\u5E8A\u53F7")
      bedid.Num=sub(pattern = toUTF8("\u95E8\u8BCA/\u4F4F\u9662\u53F7"),replacement = "",
                    sub(pattern = bedid,replacement = "",
                        x =str_extract(string = TRdata2.i,
                                       pattern = paste0(bedid,".*",toUTF8("\u95E8\u8BCA/\u4F4F\u9662\u53F7")))))
      #Hospitalid
      Hospitalid=toUTF8("\u95E8\u8BCA/\u4F4F\u9662\u53F7")
      Hospitalid.Num=sub(pattern =  paste0('[ ,]*',toUTF8("\u5E74\u9F84")),replacement = "",
                         sub(pattern = Hospitalid,replacement = "",
                             x =str_extract(string = TRdata2.i,
                                            pattern = paste0(Hospitalid,".*", toUTF8("\u5E74\u9F84")))))
      #age
      age=toUTF8("\u5E74\u9F84")
      age.Num=sub(pattern =  paste0('[ ,]*',toUTF8('\u63A5\u6536\u65F6\u95F4')),replacement = "",
                  sub(pattern = age,replacement = "",
                      x =str_extract(string = TRdata2.i,
                                     pattern = paste0(age,".*", toUTF8('\u63A5\u6536\u65F6\u95F4')))))
      #AcceptTime
      AcceptTime=toUTF8('\u63A5\u6536\u65F6\u95F4')
      AcceptTime.Num=sub(pattern =  paste0('[ ,]*',toUTF8("\u7533\u8BF7\u533B\u751F")),replacement = "",
                         sub(pattern = AcceptTime,replacement = "",
                             x =str_extract(string = TRdata2.i,
                                            pattern = paste0(AcceptTime,".*", toUTF8("\u7533\u8BF7\u533B\u751F")))))
      #ApplyDoctor
      ApplyDoctor=toUTF8("\u7533\u8BF7\u533B\u751F")
      ApplyDoctor.Num=sub(pattern =  paste0('[ ,]*',toUTF8("\u9001\u68C0\u6750\u6599")),replacement = "",
                          sub(pattern = ApplyDoctor,replacement = "",
                              x =str_extract(string = TRdata2.i,
                                             pattern = paste0(ApplyDoctor,".*", toUTF8("\u9001\u68C0\u6750\u6599")))))
      #bacterial
      bacterial=toUTF8("\u9001\u68C0\u6750\u6599")
      bacterial.Num=sub(pattern =  paste0('[ ,]*',toUTF8("\u91C7\u6837\u65F6\u95F4")),replacement = "",
                        sub(pattern = bacterial,replacement = "",
                            x =str_extract(string = TRdata2.i,
                                           pattern = paste0(bacterial,".*", toUTF8("\u91C7\u6837\u65F6\u95F4")))))
      #SamplingTime
      SamplingTime=toUTF8("\u91C7\u6837\u65F6\u95F4")
      SamplingTime.Num=gsub(pattern =  '[a-zA-Z ,]*',replacement = "",
                            sub(pattern = SamplingTime,replacement = "",
                                x =str_extract(string = TRdata2.i,
                                               pattern = paste0(SamplingTime,"[a-zA-Z0-9- ]*"))))
      #SamplingCondition
      SamplingCondition=toUTF8("\u6807\u672C\u60C5\u51B5")
      SamplingCondition.Num=sub(pattern =  toUTF8("\u8054\u7CFB\u7535\u8BDD"),replacement = "",
                                sub(pattern = SamplingCondition,replacement = "",
                                    x =str_extract(string = TRdata2.i,
                                                   pattern = paste0(SamplingCondition,".*",toUTF8("\u8054\u7CFB\u7535\u8BDD")))))
      #LymphoidRatio
      LymphoidRatio='Lymphoid'
      LymphoidRatio.Num=gsub(pattern =  '[ ,]*',replacement = "",
                            sub(pattern = LymphoidRatio,replacement = "",
                                x =str_extract(string = TRdata2.i,
                                               pattern = paste0(LymphoidRatio,"[0-9. ]*"))))

      #GransRatio
      GransRatio='Grans'
      GransRatio.Num=gsub(pattern =  '[ ,]*',replacement = "",
                         sub(pattern = GransRatio,replacement = "",
                             x =str_extract(string = TRdata2.i,
                                            pattern = paste0(GransRatio,"[0-9. ]*"))))

      #MonosRatio
      MonosRatio='Monos'
      MonosRatio.Num=gsub(pattern =  '[ ,]*',replacement = "",
                         sub(pattern = MonosRatio,replacement = "",
                             x =str_extract(string = TRdata2.i,
                                            pattern = paste0(MonosRatio,"[0-9. ]*"))))
      #CD45Weak
      CD45Weak=toUTF8("CD45\u5F31\u8868\u8FBE\u7EC6\u80DE")
      CD45Weak.Num=sub(pattern =  '[ ,]*',replacement = "",
                       sub(pattern = CD45Weak,replacement = "",
                           x =str_extract(string = TRdata2.i,
                                          pattern = paste0(CD45Weak,"[,0-9. ]*"))))
      #CD45Negative
      CD45Negative=toUTF8("CD45\u9634\u6027\u8868\u8FBE\u7EC6\u80DE")
      CD45Negative.Num=sub(pattern =  '[ ,]*',replacement = "",
                           sub(pattern = CD45Negative,replacement = "",
                               x =str_extract(string = TRdata2.i,
                                              pattern = paste0(CD45Negative,"[,0-9. ]*"))))
      #SampleViability
      SampleViability='Sample Viability'
      SampleViability.Num=gsub(pattern =  '[): ]*',replacement = "",
                               sub(pattern = SampleViability,replacement = "",
                                   x =str_extract(string = TRdata2.i,
                                                  pattern = paste0(SampleViability,"[): 0-9]*%"))))
      #CellYield
      CellYield='Cell Yield'
      CellYield.Num=gsub(pattern =  '[): ]*',replacement = "",
                         sub(pattern = CellYield,replacement = "",
                             x =str_extract(string = TRdata2.i,
                                            pattern = paste0(CellYield,"[): 0-9]*"))))
      #lymphy
      lymphy=toUTF8("\u6DCB\u5DF4\u7EC6\u80DE\u8868\u578B\u5206\u6790")
      lymphy.Num=gsub(pattern =  toUTF8("[ ,]{0,}\u7C92\u7EC6\u80DE, Grans"),replacement = "",
                      sub(pattern = 'Lymphoid[ ,.0-9]*',replacement = "",
                          x =str_extract(string = TRdata2.i,
                                         pattern = paste0('Lymphoid[ ,.0-9]*',".*",toUTF8("\u7C92\u7EC6\u80DE, Grans")))))
      #Gransphy
      Gransphy=toUTF8("\u7C92\u7EC6\u80DE\u8868\u578B\u5206\u6790")
      Gransphy.Num=gsub(pattern =  toUTF8("[ ,]{0,}\u5355\u6838\u7EC6\u80DE, Monos"),replacement = "",
                        sub(pattern = 'Grans[ ,.0-9]*',replacement = "",
                            x =str_extract(string = TRdata2.i,
                                           pattern = paste0('Grans[ ,.0-9]*',".*",toUTF8("\u5355\u6838\u7EC6\u80DE, Monos")))))
      #Monosphy
      Monosphy=toUTF8("\u5355\u6838\u7EC6\u80DE\u8868\u578B\u5206\u6790")
      Monosphy.Num=gsub(pattern =  paste0("[ ,]{0,}",CD45Weak),replacement = "",
                        sub(pattern = 'Monos[ ,.0-9]*',replacement = "",
                            x =str_extract(string = TRdata2.i,
                                           pattern = paste0('Monos[ ,.0-9]*',".*",CD45Weak))))
      #CD45weakphy
      CD45weakphy=toUTF8("CD45\u5F31\u8868\u8FBE\u7EC6\u80DE\u8868\u578B\u5206\u6790")
      CD45weakphy.Num=gsub(pattern =  paste0("[ ,]{0,}",CD45Negative),replacement = "",
                           sub(pattern = paste0(CD45Weak,'[ ,.0-9]*'),replacement = "",
                               x =str_extract(string = TRdata2.i,
                                              pattern = paste0(CD45Weak,".*",CD45Negative))))
      #CD45Negativephy
      CD45Negativephy=toUTF8("CD45\u9634\u6027\u8868\u8FBE\u7EC6\u80DE\u8868\u578B\u5206\u6790")
      CD45Negativephy.Num=gsub(pattern =  paste0("[ ,(]{0,}",toUTF8("\u6837\u672C\u6D3B\u6027"),"[ ,(]{0,}",'Sample Viability'),replacement = "",
                               sub(pattern = paste0(CD45Negative,'[ ,.0-9]*'),replacement = "",
                                   x =str_extract(string = TRdata2.i,
                                                  pattern = paste0(CD45Negative,".*",'Sample Viability'))))
      #Conclusion
      Conclusion=toUTF8("\u5206\u6790\u7ED3\u8BBA")
      Conclusion.Num=gsub(pattern =  paste0("[ ,\\(]{0,}",toUTF8("\u89E3\u91CA\u4E0E\u610F\u89C1")),replacement = "",
                          gsub(pattern = paste0('[ ,.0-9]*','Conclusion\\)[ :,]*'),replacement = "",
                               x =str_extract(string = TRdata2.i,
                                              pattern = paste0('(Conclusion)',".*",toUTF8("\u89E3\u91CA\u4E0E\u610F\u89C1")))))
      #Interpretation
      Interpretation=toUTF8("\u89E3\u91CA\u4E0E\u610F\u89C1")
      Interpretation.Num=gsub(pattern =  ".*Comments)[ :,]*",replacement = "",
                              gsub(pattern = paste0('[ ,]*',toUTF8("\u6B64\u62A5\u544A\u68C0\u6D4B\u7684CD\\(Markers")),replacement = "",
                                   x =str_extract(string = TRdata2.i,
                                                  pattern = paste0('\\(Interpretation',".*","CD\\(Markers"))))

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
                          SamplingCondition.Num,
                          LymphoidRatio.Num,
                          GransRatio.Num,
                          MonosRatio.Num,
                          CD45Weak.Num,
                          CD45Negative.Num,
                          SampleViability.Num,
                          CellYield.Num,
                          lymphy.Num,
                          Gransphy.Num,
                          Monosphy.Num,
                          CD45weakphy.Num ,
                          CD45Negativephy.Num,
                          Conclusion.Num,
                          Interpretation.Num)
      TR.datafrme=rbind(TR.datafrme,TR.datafrme.i)
      setTxtProgressBar(pb = prgbar, value = flui.i)
    }
    close(prgbar)
    colnames(TR.datafrme)= c(SampleCode,
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
                             SamplingCondition,
                             toUTF8("\u6DCB\u5DF4\u7EC6\u80DE\u6BD4\u4F8B"),
                             toUTF8("\u7C92\u7EC6\u80DE\u6BD4\u4F8B"),
                             toUTF8("\u5355\u6838\u7EC6\u80DE\u6BD4\u4F8B"),
                             CD45Weak,
                             CD45Negative,
                             toUTF8("\u6837\u672C\u6D3B\u6027"),
                             toUTF8("\u9002\u68C0\u7EC6\u80DE\u6570"),
                             lymphy,
                             Gransphy,
                             Monosphy,
                             CD45weakphy ,
                             CD45Negativephy,
                             Conclusion,
                             Interpretation)
    TR.datafrme1=data.frame(ifelse(is.na(as.matrix(TR.datafrme)),"",as.matrix(TR.datafrme)))
    colnames(TR.datafrme1)=colnames(TR.datafrme)
    TR.datafrme=TR.datafrme1
    cat( toUTF8("\u68C0\u6D4B\u9879\u76EE:"),
         "\n",
         toUTF8("\u6D41\u5F0F")
    )
    if (WriteExecl==TRUE){
      filenamewritten=Sys.time() %>%
        sub(pattern = ":",replacement = toUTF8("\u02B1"))%>%
        sub(pattern = ":",replacement = toUTF8("\u5206"))%>%
        paste0(toUTF8("\u79D2 \u6D41\u5F0F.csv"))
      write.csv(TR.datafrme,filenamewritten)
      cat("\n","\n",
          toUTF8("\u6570\u636E\u88AB\u5199\u5165Excel\u4E2D,\u6587\u4EF6\u540D\u4E3A:"),
          "\n",
          filenamewritten)
    }
    return(TR.datafrme)
}
