import Vue from 'vue'
import App from './App.vue'

Vue.config.productionTip = false

const Office = window.Office

new Vue({
  render: h => h(App)
}).$mount('#app')

Office.initialize = () => {}
