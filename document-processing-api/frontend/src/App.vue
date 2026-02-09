<template>
  <div class="min-h-screen bg-gray-50">
    <div class="container mx-auto py-8">
      <h1 class="text-3xl font-bold text-center mb-8">Document Processing Service</h1>

      <!-- Processor Selection -->
      <div class="max-w-xl mx-auto mb-6">
        <label class="block text-sm font-medium text-gray-700 mb-2">Select Processor</label>
        <select v-model="processorType" class="w-full px-3 py-2 border border-gray-300 rounded-md">
          <option value="word_to_html">Word to HTML</option>
        </select>
      </div>

      <!-- Conversion Settings -->
      <div v-if="processorType === 'word_to_html'"
           class="max-w-xl mx-auto mb-6 p-4 bg-white rounded-lg shadow-sm border border-gray-200">
        <h3 class="text-lg font-semibold mb-4 text-gray-800">Conversion Settings</h3>

        <!-- Equation Marker Style -->
        <div class="mb-4">
          <label class="block text-sm font-medium text-gray-700 mb-2">Equation Marker Style</label>
          <div class="flex space-x-2">
            <button
              @click="setMarkerPreset('none')"
              :class="['px-3 py-2 text-sm rounded-md border', markerPreset === 'none' ? 'bg-blue-500 text-white border-blue-500' : 'bg-white text-gray-700 border-gray-300 hover:bg-gray-50']"
            >None</button>
            <button
              @click="setMarkerPreset('markers')"
              :class="['px-3 py-2 text-sm rounded-md border', markerPreset === 'markers' ? 'bg-blue-500 text-white border-blue-500' : 'bg-white text-gray-700 border-gray-300 hover:bg-gray-50']"
            >MATHSTART/END</button>
            <button
              @click="setMarkerPreset('custom')"
              :class="['px-3 py-2 text-sm rounded-md border', markerPreset === 'custom' ? 'bg-blue-500 text-white border-blue-500' : 'bg-white text-gray-700 border-gray-300 hover:bg-gray-50']"
            >Custom</button>
          </div>
        </div>

        <!-- Custom Prefix/Suffix (shown when custom is selected) -->
        <div v-if="markerPreset === 'custom'" class="mb-4 grid grid-cols-2 gap-4">
          <div>
            <label class="block text-sm text-gray-600 mb-1">Inline Prefix</label>
            <input v-model="conversionConfig.inline_prefix" type="text"
                   class="w-full px-3 py-2 border border-gray-300 rounded-md text-sm">
          </div>
          <div>
            <label class="block text-sm text-gray-600 mb-1">Inline Suffix</label>
            <input v-model="conversionConfig.inline_suffix" type="text"
                   class="w-full px-3 py-2 border border-gray-300 rounded-md text-sm">
          </div>
          <div>
            <label class="block text-sm text-gray-600 mb-1">Display Prefix</label>
            <input v-model="conversionConfig.display_prefix" type="text"
                   class="w-full px-3 py-2 border border-gray-300 rounded-md text-sm">
          </div>
          <div>
            <label class="block text-sm text-gray-600 mb-1">Display Suffix</label>
            <input v-model="conversionConfig.display_suffix" type="text"
                   class="w-full px-3 py-2 border border-gray-300 rounded-md text-sm">
          </div>
        </div>

        <!-- Checkboxes -->
        <div class="space-y-3">
          <label class="flex items-center space-x-3 cursor-pointer">
            <input type="checkbox" v-model="conversionConfig.convert_shapes_to_svg"
                   class="w-4 h-4 text-blue-500 rounded">
            <span class="text-sm text-gray-700">Convert shapes to SVG</span>
          </label>

          <label class="flex items-center space-x-3 cursor-pointer">
            <input type="checkbox" v-model="conversionConfig.include_images"
                   class="w-4 h-4 text-blue-500 rounded">
            <span class="text-sm text-gray-700">Include images in HTML</span>
          </label>

          <label class="flex items-center space-x-3 cursor-pointer">
            <input type="checkbox" v-model="conversionConfig.include_mathjax"
                   class="w-4 h-4 text-blue-500 rounded">
            <span class="text-sm text-gray-700">Include MathJax library</span>
          </label>

          <label class="flex items-center space-x-3 cursor-pointer">
            <input type="checkbox" v-model="conversionConfig.rtl_direction"
                   class="w-4 h-4 text-blue-500 rounded">
            <span class="text-sm text-gray-700">RTL direction (Arabic/Hebrew)</span>
          </label>
        </div>
      </div>

      <!-- File Upload -->
      <FileUploader
        @files-selected="handleFiles"
        @upload="uploadFiles"
        :files="files"
      />

      <!-- Job Status -->
      <JobStatus
        v-if="jobId"
        :job-id="jobId"
        @completed="handleCompleted"
      />

      <!-- Results -->
      <ResultDownload
        v-if="results.length"
        :results="results"
        :job-id="completedJobId"
        @reset="resetProcessor"
      />
    </div>
  </div>
</template>

<script>
import { ref, reactive } from 'vue'
import FileUploader from './components/FileUploader.vue'
import JobStatus from './components/JobStatus.vue'
import ResultDownload from './components/ResultDownload.vue'
import axios from 'axios'
import { API_BASE_URL } from './config'

export default {
  components: {
    FileUploader,
    JobStatus,
    ResultDownload
  },
  setup() {
    const files = ref([])
    const jobId = ref(null)
    const completedJobId = ref(null)
    const results = ref([])
    const processorType = ref('word_to_html')
    const markerPreset = ref('none')

    // Conversion configuration
    const conversionConfig = reactive({
      output_format: 'latex_html',  // Default: LaTeX + MathJax
      inline_prefix: '',
      inline_suffix: '',
      display_prefix: '',
      display_suffix: '',
      convert_shapes_to_svg: false,  // Default: ignore Word shapes
      include_images: true,
      include_mathjax: true,
      rtl_direction: true
    })

    const setMarkerPreset = (preset) => {
      markerPreset.value = preset
      if (preset === 'none') {
        conversionConfig.inline_prefix = ''
        conversionConfig.inline_suffix = ''
        conversionConfig.display_prefix = ''
        conversionConfig.display_suffix = ''
      } else if (preset === 'markers') {
        conversionConfig.inline_prefix = 'MATHSTARTINLINE'
        conversionConfig.inline_suffix = 'MATHENDINLINE'
        conversionConfig.display_prefix = 'MATHSTARTDISPLAY'
        conversionConfig.display_suffix = 'MATHENDDISPLAY'
      }
      // 'custom' - keep current values
    }

    const handleFiles = (selectedFiles) => {
      files.value = selectedFiles
    }

    const uploadFiles = async () => {
      const formData = new FormData()
      files.value.forEach(file => formData.append('files', file))
      formData.append('processor_type', processorType.value)

      // Add conversion config - convert reactive proxy to plain object
      const configToSend = {
        output_format: conversionConfig.output_format,
        inline_prefix: conversionConfig.inline_prefix,
        inline_suffix: conversionConfig.inline_suffix,
        display_prefix: conversionConfig.display_prefix,
        display_suffix: conversionConfig.display_suffix,
        convert_shapes_to_svg: conversionConfig.convert_shapes_to_svg,
        include_images: conversionConfig.include_images,
        include_mathjax: conversionConfig.include_mathjax,
        rtl_direction: conversionConfig.rtl_direction
      }
      console.log('Config to send:', configToSend)
      formData.append('conversion_config', JSON.stringify(configToSend))

      try {
        const response = await axios.post(`${API_BASE_URL}/api/process`, formData)
        jobId.value = response.data.job_id
        console.log('Job started:', jobId.value)
        console.log('Full config sent:', configToSend)
        files.value = [] // Clear files after upload
      } catch (error) {
        console.error('Upload failed:', error)
        alert('Upload failed. Please check if the backend is running.')
      }
    }

    const handleCompleted = (jobResults) => {
      console.log('[App] Job completed event received')
      console.log('[App] Results:', jobResults)

      if (!jobResults || jobResults.length === 0) {
        console.warn('[App] No results received!')
        alert('Job completed but no results found. Check console for details.')
      }

      results.value = jobResults
      completedJobId.value = jobId.value
      jobId.value = null

      console.log('[App] State updated - results:', results.value.length, 'completedJobId:', completedJobId.value)
    }

    const resetProcessor = () => {
      files.value = []
      jobId.value = null
      completedJobId.value = null
      results.value = []
    }

    return {
      files,
      jobId,
      completedJobId,
      results,
      processorType,
      markerPreset,
      conversionConfig,
      setMarkerPreset,
      handleFiles,
      uploadFiles,
      handleCompleted,
      resetProcessor
    }
  }
}
</script>