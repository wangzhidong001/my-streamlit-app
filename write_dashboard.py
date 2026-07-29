import os

content = '''<template>
  <div class="dashboard">
    <el-header class="header">
      <div class="logo">AI 数据分析平台</div>
      <div class="user-info">
        <span>{{ userStore.userInfo?.username }} ({{ userStore.userInfo?.role }})</span>
        <el-button type="danger" size="small" @click="handleLogout">退出</el-button>
      </div>
    </el-header>

    <el-main>
      <el-row :gutter="20" class="summary-row">
        <el-col :span="6">
          <el-card>
            <div class="summary-label">总销售额</div>
            <div class="summary-value">{{ formatNumber(summary.total_sales) }}</div>
          </el-card>
        </el-col>
        <el-col :span="6">
          <el-card>
            <div class="summary-label">总利润</div>
            <div class="summary-value">{{ formatNumber(summary.total_profit) }}</div>
          </el-card>
        </el-col>
        <el-col :span="6">
          <el-card>
            <div class="summary-label">利润率</div>
            <div class="summary-value">{{ summary.profit_rate }}%</div>
          </el-card>
        </el-col>
        <el-col :span="6">
          <el-card>
            <div class="summary-label">订单数量</div>
            <div class="summary-value">{{ summary.order_count }}</div>
          </el-card>
        </el-col>
      </el-row>

      <el-row :gutter="20" class="chart-row">
        <el-col :span="12">
          <el-card>
            <template #header>月度销售趋势</template>
            <div ref="lineChart" style="height: 350px"></div>
          </el-card>
        </el-col>
        <el-col :span="12">
          <el-card>
            <template #header>区域销售分布</template>
            <div ref="pieChart" style="height: 350px"></div>
          </el-card>
        </el-col>
      </el-row>

      <el-card class="drill-card">
        <template #header>
          <div class="card-header">
            <span>数据下钻分析</span>
            <el-breadcrumb separator=">">
              <el-breadcrumb-item
                v-for="(item, index) in breadcrumb"
                :key="index"
                @click="drillUp(index)"
                :class="{ clickable: index < breadcrumb.length - 1 }"
              >
                {{ item.label }}
              </el-breadcrumb-item>
            </el-breadcrumb>
          </div>
        </template>

        <el-table :data="drillData" style="width: 100%" @row-click="handleDrillDown">
          <el-table-column prop="dimension" label="维度" />
          <el-table-column prop="sales" label="销售额" :formatter="(row) => formatNumber(row.sales)" />
          <el-table-column prop="profit" label="利润" :formatter="(row) => formatNumber(row.profit)" />
          <el-table-column prop="order_count" label="订单数" />
          <el-table-column label="操作">
            <template #default="{ row }">
              <el-button
                v-if="currentLevel !== 'detail'"
                type="primary"
                size="small"
                @click.stop="handleDrillDown(row)"
              >
                下钻
              </el-button>
            </template>
          </el-table-column>
        </el-table>
      </el-card>

      <el-card class="simulate-card">
        <template #header>利润模拟测算</template>

        <el-row :gutter="40">
          <el-col :span="8">
            <div class="param-item">
              <label>价格调整: {{ (simulateParams.price_adjustment * 100).toFixed(0) }}%</label>
              <el-slider
                v-model="simulateParams.price_adjustment"
                :min="-0.3"
                :max="0.3"
                :step="0.01"
                @change="runSimulation"
              />
            </div>
            <div class="param-item">
              <label>成本调整: {{ (simulateParams.cost_adjustment * 100).toFixed(0) }}%</label>
              <el-slider
                v-model="simulateParams.cost_adjustment"
                :min="-0.3"
                :max="0.3"
                :step="0.01"
                @change="runSimulation"
              />
            </div>
            <div class="param-item">
              <label>销量调整: {{ (simulateParams.volume_adjustment * 100).toFixed(0) }}%</label>
              <el-slider
                v-model="simulateParams.volume_adjustment"
                :min="-0.5"
                :max="0.5"
                :step="0.01"
                @change="runSimulation"
              />
            </div>
            <el-button type="primary" @click="resetSimulation">重置参数</el-button>
          </el-col>

          <el-col :span="16">
            <el-row :gutter="20">
              <el-col :span="8">
                <div class="compare-box">
                  <div class="compare-title">原始利润</div>
                  <div class="compare-value">{{ formatNumber(simulateResult.base?.profit) }}</div>
                </div>
              </el-col>
              <el-col :span="8">
                <div class="compare-box">
                  <div class="compare-title">模拟利润</div>
                  <div class="compare-value highlight">{{ formatNumber(simulateResult.simulated?.profit) }}</div>
                </div>
              </el-col>
              <el-col :span="8">
                <div class="compare-box">
                  <div class="compare-title">利润变化</div>
                  <div class="compare-value" :class="{ up: profitChange > 0, down: profitChange < 0 }">
                    {{ profitChange > 0 ? '+' : '' }}{{ formatNumber(profitChange) }}
                    ({{ profitChangeRate }}%)
                  </div>
                </div>
              </el-col>
            </el-row>

            <div ref="simulateChart" style="height: 280px; margin-top: 20px"></div>
          </el-col>
        </el-row>
      </el-card>
    </el-main>
  </div>
</template>

<script setup>
import { ref, reactive, onMounted, onUnmounted, computed, nextTick } from 'vue'
import { useRouter } from 'vue-router'
import * as echarts from 'echarts'
import request from '@/utils/request'
import { useUserStore } from '@/stores/user'
import { ElMessage } from 'element-plus'

const router = useRouter()
const userStore = useUserStore()

const summary = reactive({
  total_sales: 0,
  total_profit: 0,
  profit_rate: 0,
  order_count: 0
})

const lineChart = ref(null)
const pieChart = ref(null)
const simulateChart = ref(null)
let lineChartInst = null
let pieChartInst = null
let simulateChartInst = null

const drillData = ref([])
const currentLevel = ref('year')
const breadcrumb = ref([{ label: '年度', level: 'year', value: null }])

const simulateParams = reactive({
  price_adjustment: 0,
  cost_adjustment: 0,
  volume_adjustment: 0
})
const simulateResult = reactive({ base: {}, simulated: {}, change: {} })

const profitChange = computed(() => simulateResult.change?.profit_change || 0)
const profitChangeRate = computed(() => simulateResult.change?.profit_change_rate || 0)

function formatNumber(val) {
  if (val === undefined || val === null) return '-'
  return val.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
}

async function fetchSummary() {
  const data = await request.get('/data/summary')
  Object.assign(summary, data)
}

async function initLineChart() {
  const data = await request.get('/chart/monthly-trend')
  lineChartInst = echarts.init(lineChart.value)
  lineChartInst.setOption({
    tooltip: { trigger: 'axis' },
    legend: { data: data.series.map(s => s.name) },
    xAxis: { type: 'category', data: data.xAxis },
    yAxis: { type: 'value' },
    series: data.series.map(s => ({
      ...s,
      smooth: true,
      areaStyle: { opacity: 0.1 }
    }))
  })
}

async function initPieChart() {
  const data = await request.get('/chart/region-distribution')
  pieChartInst = echarts.init(pieChart.value)
  pieChartInst.setOption({
    tooltip: { trigger: 'item', formatter: '{b}: {c} ({d}%)' },
    series: [{
      type: 'pie',
      radius: ['40%', '70%'],
      data: data.data,
      emphasis: {
        itemStyle: { shadowBlur: 10, shadowOffsetX: 0, shadowColor: 'rgba(0,0,0,0.5)' }
      }
    }]
  })
}

async function loadDrillData(level, parentValue = null) {
  currentLevel.value = level
  let url = `/drill/${level}`
  const params = {}

  if (level === 'quarter' && parentValue) params.year = parentValue
  if (level === 'month') {
    const parent = breadcrumb.value[breadcrumb.value.length - 2]
    if (parent?.level === 'year') params.year = parent.value
    if (parent?.level === 'quarter') {
      params.year = breadcrumb.value.find(b => b.level === 'year')?.value
      params.quarter = parent.value
    }
  }
  if (level === 'detail' && parentValue) params.month = parentValue

  const queryString = new URLSearchParams(params).toString()
  if (queryString) url += '?' + queryString

  const res = await request.get(url)
  drillData.value = res.items
}

function handleDrillDown(row) {
  const levelMap = { year: 'quarter', quarter: 'month', month: 'detail' }
  const nextLevel = levelMap[currentLevel.value]
  if (!nextLevel) return

  breadcrumb.value.push({
    label: row.dimension,
    level: nextLevel,
    value: row.dimension
  })
  loadDrillData(nextLevel, row.dimension)
}

function drillUp(index) {
  if (index >= breadcrumb.value.length - 1) return
  breadcrumb.value = breadcrumb.value.slice(0, index + 1)
  const item = breadcrumb.value[index]
  loadDrillData(item.level, item.value)
}

async function runSimulation() {
  const res = await request.post('/simulate/profit', {
    price_adjustment: simulateParams.price_adjustment,
    cost_adjustment: simulateParams.cost_adjustment,
    volume_adjustment: simulateParams.volume_adjustment
  })
  Object.assign(simulateResult, res)
  updateSimulateChart()
}

function updateSimulateChart() {
  if (!simulateChartInst) {
    simulateChartInst = echarts.init(simulateChart.value)
  }
  simulateChartInst.setOption({
    tooltip: { trigger: 'axis' },
    xAxis: { type: 'category', data: ['原始', '模拟'] },
    yAxis: { type: 'value' },
    series: [
      {
        name: '销售额',
        type: 'bar',
        data: [simulateResult.base?.sales || 0, simulateResult.simulated?.sales || 0]
      },
      {
        name: '利润',
        type: 'bar',
        data: [simulateResult.base?.profit || 0, simulateResult.simulated?.profit || 0]
      }
    ]
  })
}

function resetSimulation() {
  simulateParams.price_adjustment = 0
  simulateParams.cost_adjustment = 0
  simulateParams.volume_adjustment = 0
  runSimulation()
}

function handleLogout() {
  userStore.logout()
  ElMessage.success('已退出登录')
  router.push('/login')
}

onMounted(async () => {
  await fetchSummary()
  await initLineChart()
  await initPieChart()
  await loadDrillData('year')
  await runSimulation()
})

function handleResize() {
  lineChartInst?.resize()
  pieChartInst?.resize()
  simulateChartInst?.resize()
}
window.addEventListener('resize', handleResize)
onUnmounted(() => {
  window.removeEventListener('resize', handleResize)
  lineChartInst?.dispose()
  pieChartInst?.dispose()
  simulateChartInst?.dispose()
})
</script>

<style scoped>
.dashboard { min-height: 100vh; background: #f5f7fa; }
.header {
  background: #fff;
  display: flex;
  align-items: center;
  justify-content: space-between;
  box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}
.logo { font-size: 20px; font-weight: bold; color: #409eff; }
.user-info { display: flex; align-items: center; gap: 16px; }

.summary-row { margin-bottom: 20px; }
.summary-label { font-size: 14px; color: #999; margin-bottom: 8px; }
.summary-value { font-size: 24px; font-weight: bold; color: #333; }

.chart-row { margin-bottom: 20px; }

.drill-card { margin-bottom: 20px; }
.card-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
}
.clickable { cursor: pointer; color: #409eff; }
.clickable:hover { text-decoration: underline; }

.simulate-card { margin-bottom: 20px; }
.param-item { margin-bottom: 20px; }
.param-item label { display: block; margin-bottom: 8px; font-size: 14px; color: #666; }

.compare-box {
  background: #f5f7fa;
  border-radius: 8px;
  padding: 20px;
  text-align: center;
}
.compare-title { font-size: 14px; color: #999; margin-bottom: 8px; }
.compare-value { font-size: 20px; font-weight: bold; color: #333; }
.compare-value.highlight { color: #409eff; }
.compare-value.up { color: #67c23a; }
.compare-value.down { color: #f56c6c; }
</style>
'''

filepath = r'C:\Users\ruijie\Desktop\ai-data-app\frontend\src\views\DashboardView.vue'
with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)
print(f'DashboardView.vue created successfully ({len(content)} bytes)')
