<template>
  <div class="hello">
    <div class="global">
        <div>
          <el-button style="margin-left:90px;" type="primary" @click="onFileSelect">文件选择</el-button>
        </div>
        <hr>
        <div class="divItem">
            <span>编码位置：</span>
            <input type="text" :value="itemLocation" name="itemLocation">
        </div>
        <div class="divItem">
          <span>请指定上方参考值：</span>
          <input type="text" :value="frontValue" name="frontValue">
        </div>
        <div class="divItem">
          <span>请指定下方参考值：</span>
          <input type="text" :value="endValue" name="endValue">
        </div>
        <div class="divItem">如果是最后一个项目，请输入”末尾“</div>
        <hr>
        <div class="divItem">
            <el-button style="margin-left:90px;" type="primary" @click="onSure">确定</el-button>
            <el-button style="margin-left:90px;" type="primary" @click="onCancel">取消</el-button>
        </div>
    </div>
  </div>
</template>

<script>
import tool from './js/tool.js'
import axios from 'axios'
export default {
  name: 'Dialog',
  data(){
      return {
          DemoSpan : '',
          docName: '',
          itemLocation: 'H4',
          // frontValue: '项目支出明细测算',
          // endValue: '项目采购',
          frontValue: '21',
          endValue: '34',
          filePathList: []
      }
  },
  methods:{
      onSure(){
          let result = tool.extractTemp(this.filePathList, this.itemLocation, this.frontValue, this.endValue)
          console.log(result)
          alert(result)
      },
      onCancel(){
          close()
      },
      onFileSelect(){
        this.filePathList = wps.EtApplication().GetOpenFilename("表格文件(*.xls;*.xlsx), *.xls;*.xlsx", 1, "表格文件选择", "确定;取消", true)
        console.log(this.filePathList)
      }
  },
  mounted() {
      axios.get('/.debugTemp/NotifyDemoUrl').then((res) => {
          this.DemoSpan = res.data;
      });
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
.global{
    font-size: 15px;
    min-height: 95%;
    }
.divItem{
    margin-left: 5px;
    margin-bottom: 18px;
    font-size: 15px;
    word-wrap:break-word;
    }
</style>
