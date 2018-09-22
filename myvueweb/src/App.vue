<template>
  <div id="app">
    <img alt="Vue logo"
      src="./assets/logo.png">
    <button @click="write">写入</button>
    <button @click="read">读取</button>
    <button @click="deleteOne">删除</button>
    <button @click="clear">清空</button>

    <div>{{content}}</div>
    <div>{{numberFormat}}</div>
  </div>
</template>

<script>
import HelloWorld from './components/HelloWorld.vue'

export default {
  name: 'app',
  components: {
    HelloWorld
  },
  data: function() {
    return {
      numberFormat: null,
      content: null
    }
  },
  methods: {
    /**
     * 写入
     */
    write: function() {
      Excel.run(function(context) {
        //获取指定名字的工作薄sheet
        // var sheet = context.workbook.worksheets.getItem('sheetName')
        //获取当前活动工作薄
        var sheet = context.workbook.worksheets.getActiveWorksheet()
        var data = [
          ['Product', 'Qty', 'Unit Price', 'Total Price'],
          ['Almonds', '2', '7.5', '15'],
          ['Coffee', '1', '34.5', '34.5'],
          ['Chocolate', '5', '9.56', '47.8'],
          ['', '', '', '97.3']
        ]
        //此二维数组的长度要和数据的保持一致，否则无效
        var formats = [
          ['@', '@', '@', '@'], //设置格式为文本
          ['0.00', '0.00', '0.00', '0.00'],
          ['0.00', '0.00', '0.00', '0.00'],
          ['0.00', '0.00', '0.00', '0.00'],
          ['0.00', '0.00', '0.00', '0.00']
        ]

        var range = sheet.getRange('A1:D5')
        //选中该区域
        range.select()
        // 设置背景色和字体
        range.format.fill.color = '#4472C4'
        range.format.font.color = 'white'
        //设置区域的格式
        range.numberFormat = formats
        //表示加载values属性，如果不加载在下面是不可以使用的
        range.load('values')

        return context.sync().then(function() {
          //写入方法必须在该方法内执行才有效
          range.values = data
        })
      }).catch(_this.errorHandler)
    },
    read: function() {
      let _this = this
      Excel.run(function(context) {
        //获取指定名字的工作薄sheet
        // var sheet = context.workbook.worksheets.getItem('sheetName')
        // 获取当前选中的单元格
        var range = context.workbook.getSelectedRange()
        //获取当前选中的单元格
        //表示加载以下属性，如果不加载在下面是不可以使用的
        range.load('values')
        range.load('address')
        range.load('formulas')
        range.load('text')

        return context.sync().then(function() {
          //写入方法必须在该方法内执行才有效
          _this.content = {
            values: range.values,
            formulas: range.formulas,
            address: range.address,
            texts: range.text
          }
          console.log(_this.content)
        })
      }).catch(_this.errorHandler)
    },
    deleteOne: function() {
      let _this = this
      Excel.run(function(context) {
        //获取指定名字的工作薄sheet
        // var sheet = context.workbook.worksheets.getItem('sheetName')
        // 获取当前选中的单元格
        var sheet = context.workbook.worksheets.getActiveWorksheet()
        var range = sheet.getRange('A2:D2')
        range.delete(Excel.DeleteShiftDirection.up)
        //提交操作
        return context.sync()
      }).catch(_this.errorHandler)
    },
    clear: function() {
      let _this = this
      Excel.run(function(context) {
        //获取指定名字的工作薄sheet
        // var sheet = context.workbook.worksheets.getItem('sheetName')
        // 获取当前选中的单元格
        var sheet = context.workbook.worksheets.getActiveWorksheet()
        var range = sheet.getRange('A1:D4')
        range.clear()
        //提交操作
        return context.sync()
      }).catch(_this.errorHandler)
    },
    /**
     * 操作excel的方法出现异常时调用的方法
     */
    errorHandler: function(ex) {
      alert(ex)
    }
  }
}
</script>

<style>
#app {
  font-family: 'Avenir', Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
</style>
