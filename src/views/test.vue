<template>
  <basic-container>
    <el-button type="primary" size="small"
               @click="downloadPPT()">下载PPT</el-button>
  </basic-container>
</template>

<script>
export default {
  data() {
    return {
      
    }
  },
  methods: {
    downloadPPT(){
      var pptx = new PptxGenJS();
				pptx.layout = 'LAYOUT_WIDE';
				//=====================================模板===================================
				pptx.defineSlideMaster({
				  title: 'MASTER_SLIDE',
				  bkgd:  'FFFFFF',
				  objects: [
				    { 'text':  { text:'问题：', options:{ x:0.2, y:0.5, fontSize: 19} } },
				    { 'image': { x:10.1, y:0.04, w:2.8, h:0.7, path:'http://127.0.0.1:8848/PPTExport/demo/index.html' } },
					{ 'rect':  { x: 0.0, y:0.8, w:'100%', h:0.02, fill:'f3f309' } },
					{ 'text':  { text:'零件供应商：', options:{ x:0.3, y:0.9, fontSize: 11} } },
					{ 'rect':  { x: 1.3, y:0.94, w: 2, h:0.23, fill:'f3f309' } },
					{ 'text':  { text:'发布日期：', options:{ x:3.5, y:0.9, fontSize: 11} } },
					{ 'rect':  { x: 4.3, y:0.94, w: 2, h:0.23, fill:'f3f309' } },
					{ 'text':  { text:'问题数量：', options:{ x:0.3, y:1.2, fontSize: 11} } },
					{ 'rect':  { x: 1.1, y:1.24, w: 2, h:0.23, fill:'f3f309' } },
					{ 'text':  { text:'问题来源：', options:{ x:3.5, y:1.2, fontSize: 11} } },
					{ 'rect':  { x: 4.3, y:1.24, w: 2, h:0.23, fill:'f3f309' } },
					{ 'text':  { text:'问题描述：', options:{ x:0.3, y:1.5, fontSize: 11} } },
				// ===================================内容区域======================================
					{ 'placeholder': {
							options: { name:'aaa', type:'body', x:0.3, y:1.8, w:6.3, h:1.5, line:'f3f309', lineSize: 0.02, fontSize: 11 },
							text: '问题描述信息'
						}
					},
					{ 'text':  { text:'照片：', options:{ x:0.3, y:3.3, fontSize: 11} } },
					{ 'text':  { text:'阶段：', options:{ x:2.3, y:3.3, fontSize: 11} } },
					{ 'rect':  { x: 3.3, y:3.34, w: 2, h:0.23, fill:'f3f309' } },
					{ 'placeholder': {
							options: { name:'bbb', type:'body', x:0.3, y:3.6, w:6.3, h:2, line:'f3f309', lineSize: 0.02, fontSize: 11 },
							text: '照片信息'
						}
					},
					{ 'rect':  { x: 0.0, y:6, w:'100%', h:0.02, fill:'f3f309' } }
					
				  ]
				  // slideNumber: { x:0.3, y:'90%' }
				});
				
				var slide = pptx.addSlide('MASTER_SLIDE');
				slide.addText('How To Create PowerPoint Presentations with JavaScript', { x:0.5, y:3.7, fontSize:18, color: pptx.SchemeColor.text1 });
				slide.addText('啛啛喳喳错错错错错错错错错错错错错错错错错错错错错错错错错错啛啛喳喳错错错错错错错错错错错错错错错错错错错错错啛啛喳喳错错错错错错!', { placeholder:'aaa' });
				// ====================================添加图片========================
        slide.addImage({path: 'https://mp.baomidou.com/img/relationship-with-mybatis.png', placeholder:'bbb'})
				
				pptx.writeFile();
    }
  }
};
</script>

<style>
</style>
