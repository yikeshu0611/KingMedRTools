#' To clear chromosome data from King Med test to a dataframe
#'
#' @param filename the name of chromosome data text
#' @param WriteExecl A logical object. TRUE as default to write excel
#'
#' @return a dataframe
#' @export
#'
#' @examples KingMed_Chromosome("data.txt")
KingMed_Chromosome<-function(filename,WriteExecl=TRUE){
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
    gsub(pattern = toUTF8("\u534F\u52A9\u8BCA, \u65AD"),replacement = toUTF8("\u534F\u52A9\u8BCA\u65AD"))

  #split
  tongren=toUTF8('\u4E0A\u6D77\u4EA4\u901A\u5927\u5B66\u533B\u5B66\u9662\u9644\u5C5E\u540C\u4EC1\u533B\u9662')
  TRdata1=strsplit(x = TRdata,split = tongren)[[1]]
  #get chromosome data
  dataget=c()
  chromosomeTitle=toUTF8("\u67D3\u8272\u4F53\u6838\u578B\u5206\u6790\u62A5\u544A\u5355")
  for (data.i in 1:length(TRdata1)) {
    if (suppressWarnings(grepl(chromosomeTitle,TRdata1[data.i]))){
      dataget=c(dataget,data.i)
    }
  }
  TRdata1.2=TRdata1[dataget]

  #get data page
  pageget=c()
  for (datapage.i in 1:length(TRdata1)) {
    if (grepl(toUTF8("\u6838\u578B:"),TRdata1[datapage.i])){
      pageget=c(pageget,datapage.i)
    }
  }
  TRdata2=TRdata1[pageget]

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
    PatientName=toUTF8("\u60A3\u8005\u59D3\u540D")
    PatientName.Num=sub(pattern = toUTF8("\u793E\u4F1A\u6027\u522B"),replacement = "",
                        sub(pattern = PatientName,replacement = "",
                            x =str_extract(string = TRdata2.i,
                                           pattern = paste0(PatientName,".*",toUTF8("\u793E\u4F1A\u6027\u522B")))))
    #Department
    Department=toUTF8("\u79D1\u5BA4")
    Department.Num=sub(pattern = toUTF8("\u4F4F\u9662/\u95E8\u8BCA\u53F7"),replacement = "",
                       sub(pattern = Department,replacement = "",
                           x =str_extract(string = TRdata2.i,
                                          pattern = paste0(Department,".*",toUTF8("\u4F4F\u9662/\u95E8\u8BCA\u53F7")))))
    #testid
    testid=toUTF8("\u5B9E\u9A8C\u53F7")
    testid.Num=gsub(pattern = " *",replacement = "",
                   sub(pattern = testid,replacement = "",
                       x =str_extract(string = TRdata2.i,
                                      pattern = paste0(testid,"[a-zA-Z0-9 ]*"))))
    #sex
    sex=toUTF8("\u6027\u522B")
    sex.Num=sub(pattern = toUTF8("\u5E74\u9F84"),replacement = "",
                sub(pattern = toUTF8("\u793E\u4F1A\u6027\u522B"),replacement = "",
                    x =str_extract(string = TRdata2.i,
                                   pattern = paste0(toUTF8("\u793E\u4F1A\u6027\u522B"),".*",toUTF8("\u5E74\u9F84")))))
    #bedid
    bedid=toUTF8("\u623F/\u5E8A\u53F7")
    bedid.Num=sub(pattern = paste0(" {0,}",toUTF8("\u7533\u8BF7\u533B\u751F")),replacement = "",
                  sub(pattern = bedid,replacement = "",
                      x =str_extract(string = TRdata2.i,
                                     pattern = paste0(bedid,".*",toUTF8("\u7533\u8BF7\u533B\u751F")))))
    #Hospitalid
    Hospitalid=toUTF8("\u95E8\u8BCA/\u4F4F\u9662\u53F7")
    Hospitalid.Num=gsub(pattern =  paste0('[ ,]*',toUTF8("\u5BA4\u623F/\u5E8A\u53F7")),replacement = "",
                       sub(pattern = toUTF8("\u4F4F\u9662/\u95E8\u8BCA\u53F7"),replacement = "",
                           x =str_extract(string = TRdata2.i,
                                          pattern = paste0(toUTF8("\u4F4F\u9662/\u95E8\u8BCA\u53F7"),".*", toUTF8("\u5BA4\u623F/\u5E8A\u53F7")))))
    #age
    age=toUTF8("\u5E74\u9F84")
    age.Num=sub(pattern =  paste0('[ ,]*',toUTF8('\u533B\u9662\u6807\u8BC6')),replacement = "",
                sub(pattern = age,replacement = "",
                    x =str_extract(string = TRdata2.i,
                                   pattern = paste0(age,".*", toUTF8('\u533B\u9662\u6807\u8BC6')))))
    #AcceptTime
    AcceptTime=toUTF8('\u63A5\u6536\u65F6\u95F4')
    AcceptTime.Num=sub(pattern =  paste0('[ ,]*',toUTF8("\u9879\u76EE\u540D\u79F0")),replacement = "",
                       sub(pattern = AcceptTime,replacement = "",
                           x =str_extract(string = TRdata2.i,
                                          pattern = paste0(AcceptTime,".*", toUTF8("\u9879\u76EE\u540D\u79F0")))))
    #ApplyDoctor
    ApplyDoctor=toUTF8("\u7533\u8BF7\u533B\u751F")
    ApplyDoctor.Num=sub(pattern =  paste0('[ ,]*',toUTF8("\u8054\u7CFB\u7535\u8BDD")),replacement = "",
                        sub(pattern = ApplyDoctor,replacement = "",
                            x =str_extract(string = TRdata2.i,
                                           pattern = paste0(ApplyDoctor,".*", toUTF8("\u8054\u7CFB\u7535\u8BDD")))))
    #bacterial
    bacterial=toUTF8("\u9001\u68C0\u6750\u6599")
    bacterial.Num=sub(pattern =  paste0('[ ,]*',toUTF8("\u6807\u672C\u72B6\u6001")),replacement = "",
                      sub(pattern = toUTF8("\u6807\u672C\u7C7B\u578B"),replacement = "",
                          x =str_extract(string = TRdata2.i,
                                         pattern = paste0(toUTF8("\u6807\u672C\u7C7B\u578B"),".*", toUTF8("\u6807\u672C\u72B6\u6001")))))
    #SamplingTime
    SamplingTime=toUTF8("\u91C7\u6837\u65F6\u95F4")
    SamplingTime.Num=gsub(pattern =  '[a-zA-Z ,]*',replacement = "",
                          sub(pattern = SamplingTime,replacement = "",
                              x =str_extract(string = TRdata2.i,
                                             pattern = paste0(SamplingTime,"[a-zA-Z0-9- ]*"))))
    #SamplingCondition
    SamplingCondition=toUTF8("\u6807\u672C\u60C5\u51B5")
    SamplingCondition.Num=gsub(pattern =  paste0("[;, ]*",toUTF8("\u5236\u7247\u4EBA")),replacement = "",
                              sub(pattern = toUTF8("\u6807\u672C\u72B6\u6001"),replacement = "",
                                  x =str_extract(string = TRdata2.i,
                                                 pattern = paste0(toUTF8("\u6807\u672C\u72B6\u6001"),".*",toUTF8("\u5236\u7247\u4EBA")))))
    #planting time
    PlantingTime=toUTF8("\u63A5\u79CD\u65F6\u95F4")
    PlantingTime.Num=gsub(pattern = paste0('[, ]*',toUTF8("\u6807\u672C\u7C7B\u578B")),replacement = "",
                          x = gsub(pattern = PlantingTime,replacement = "",
                                   x = str_extract(string = TRdata2.i,
                                                   pattern = paste0(PlantingTime,".*",toUTF8("\u6807\u672C\u7C7B\u578B")))) )
    #Method
    Method=toUTF8("\u663E\u5E26\u65B9\u6CD5")
    Method.Num=gsub(pattern = paste0('[, ]*',toUTF8("\u6761\u5E26\u6C34\u5E73")),replacement = "",
                    x = gsub(pattern = Method,replacement = "",
                             x = str_extract(string = TRdata2.i,
                                             pattern = paste0(Method,".*",toUTF8("\u6761\u5E26\u6C34\u5E73")))) )

    #barlevel
    BarLevel=toUTF8("\u6761\u5E26\u6C34\u5E73")
    BarLevel.Num=gsub(pattern = paste0('[, ]*',toUTF8("\u5206\u6790\u7EC6\u80DE\u6570")),replacement = "",
                      x = gsub(pattern = BarLevel,replacement = "",
                               x = str_extract(string = TRdata2.i,
                                               pattern = paste0(BarLevel,".*",toUTF8("\u5206\u6790\u7EC6\u80DE\u6570")))) )
    #CellCount
    CellCount=toUTF8("\u5206\u6790\u7EC6\u80DE\u6570")
    CellCount.Num=gsub(pattern = paste0('[, ]*',toUTF8("\u539F\u59CB\u56FE\u50CF")),replacement = "",
                       x = gsub(pattern = CellCount,replacement = "",
                                x = str_extract(string = TRdata2.i,
                                                pattern = paste0(CellCount,".*",toUTF8("\u539F\u59CB\u56FE\u50CF")))) )
    #NeuclearType
    NeuclearType=toUTF8("\u6838\u578B")
    NeuclearType.Num=gsub(pattern = paste0('[, ]*',toUTF8("\u7ED3\u679C\u89E3\u91CA")),replacement = "",
                          x = gsub(pattern = paste0(NeuclearType,":"),replacement = "",
                                   x = str_extract(string = TRdata2.i,
                                                   pattern = paste0(NeuclearType,":",".*",toUTF8("\u7ED3\u679C\u89E3\u91CA")))) )
    #Interpretation
    Interpretation=toUTF8("\u7ED3\u679C\u89E3\u91CA")
    Interpretation.Num=gsub(pattern = paste0('[a-z_, ]*',toUTF8("\u672C\u68C0\u6D4B\u4EC5\u5BF9\u6765\u6837\u8D1F\u8D23")),replacement = "",
                            x = gsub(pattern = Interpretation,replacement = "",
                                     x = str_extract(string = TRdata2.i,
                                                     pattern = paste0(Interpretation,".*",toUTF8("\u672C\u68C0\u6D4B\u4EC5\u5BF9\u6765\u6837\u8D1F\u8D23")))) )

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
                    PlantingTime.Num,
                    Method.Num,
                    BarLevel.Num,
                    CellCount.Num,
                    NeuclearType.Num,
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
      PlantingTime,
      Method,
      BarLevel,
      CellCount,
      NeuclearType,
      Interpretation)
  TR.datafrme1=data.frame(ifelse(is.na(as.matrix(TR.datafrme)),"",as.matrix(TR.datafrme)))
  colnames(TR.datafrme1)=colnames(TR.datafrme)
  TR.datafrme=TR.datafrme1
  cat( toUTF8("\u68C0\u6D4B\u9879\u76EE:"),
       "\n",
       toUTF8("\u67D3\u8272\u4F53")
  )
  if (WriteExecl==TRUE){
    filenamewritten=Sys.time() %>%
      sub(pattern = ":",replacement = toUTF8("\u02B1"))%>%
      sub(pattern = ":",replacement = toUTF8("\u5206"))%>%
      paste0(toUTF8("\u79D2 \u67D3\u8272\u4F53.csv"))
    write.csv(TR.datafrme,filenamewritten)
    cat("\n","\n",
        toUTF8("\u6570\u636E\u88AB\u5199\u5165Excel\u4E2D,\u6587\u4EF6\u540D\u4E3A:"),
        "\n",
        filenamewritten)
  }
  return(TR.datafrme)
}
