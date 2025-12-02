<template>
  <div class="max-w-xl mx-auto mt-6 p-6 bg-white rounded-lg shadow-md">
    <h3 class="text-lg font-semibold mb-4">Results Ready</h3>

    <!-- Job expired warning -->
    <div v-if="jobExpired" class="mb-4 p-3 bg-yellow-100 border border-yellow-400 text-yellow-700 rounded">
      Job has expired. Please click "Process More Documents" below to process a new document.
    </div>

    <!-- Add refresh button -->
    <button
      v-if="!jobExpired"
      @click="refreshResults"
      class="mb-3 px-3 py-1 bg-gray-100 text-gray-600 text-sm rounded hover:bg-gray-200"
    >
      Refresh Results
    </button>

    <div v-if="!jobExpired" class="space-y-3">
      <div
        v-for="(result, index) in currentResults"
        :key="index"
        class="flex justify-between items-center p-3 bg-gray-50 rounded-lg"
      >
        <div class="flex items-center space-x-3">
          <svg class="h-5 w-5 text-green-500" fill="currentColor" viewBox="0 0 20 20">
            <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clip-rule="evenodd" />
          </svg>
          <div>
            <p class="font-medium">{{ result.filename }}</p>
            <p class="text-sm text-gray-600">
              {{ formatFileSize(result.size) }}
              <span v-if="result.type === 'application/zip'" class="ml-2 text-blue-600">
                (Multiple files packaged)
              </span>
            </p>
          </div>
        </div>

        <div class="flex items-center space-x-2">
          <span v-if="result.error" class="text-xs text-red-500">{{ result.error }}</span>
          <button
            v-else
            @click="downloadFileWithRefresh(index)"
            class="px-3 py-1 bg-blue-500 text-white text-sm rounded hover:bg-blue-600 transition-colors"
          >
            {{ result.type === 'application/zip' ? 'Download ZIP' : 'Download' }}
          </button>
        </div>
      </div>
    </div>

    <!-- MathJax Script Section (hide when expired) -->
    <div v-if="!jobExpired" class="mt-6 border-t pt-4">
      <h4 class="text-sm font-semibold text-gray-700 mb-2">MathJax Script (add to HTML head):</h4>
      <div class="relative">
        <textarea
          ref="mathjaxTextarea"
          :value="mathjaxScript"
          class="w-full h-32 p-3 border border-gray-300 rounded-md font-mono text-xs bg-gray-50"
          dir="ltr"
          readonly
        ></textarea>
        <button
          @click="copyMathjaxScript"
          class="absolute top-2 right-2 px-2 py-1 bg-green-500 text-white text-xs rounded hover:bg-green-600"
        >
          {{ copySuccess ? 'Copied!' : 'Copy' }}
        </button>
      </div>
    </div>

    <button
      @click="downloadAll"
      v-if="!jobExpired && currentResults.length > 1 && !hasErrors"
      class="mt-4 w-full px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600 transition-colors font-semibold"
    >
      Download All as ZIP
    </button>

    <button
      @click="$emit('reset')"
      class="mt-3 w-full px-4 py-2 bg-gray-200 text-gray-700 rounded hover:bg-gray-300 transition-colors"
    >
      Process More Documents
    </button>
  </div>
</template>

<script>
import { ref, computed, onMounted } from 'vue'
import axios from 'axios'
import { API_BASE_URL } from '../config'

export default {
  props: ['results', 'jobId'],
  emits: ['reset'],
  setup(props, { emit }) {
    // Use local state for results that can be refreshed
    const currentResults = ref(props.results)
    const copySuccess = ref(false)
    const mathjaxTextarea = ref(null)
    const jobExpired = ref(false)

    // MathJax script for rendering equations
    const mathjaxScript = `<script>
  window.MathJax = {
    tex: {
      inlineMath: [['\\\\(', '\\\\)']],
      displayMath: [['\\\\[', '\\\\]']]
    },
    svg: { fontCache: 'global' }
  };
<\/script>
<script src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml.js" async><\/script>`

    const hasErrors = computed(() => {
      return currentResults.value.some(r => r.error)
    })

    // FORMAT FILE SIZE FUNCTION
    const formatFileSize = (bytes) => {
      if (!bytes || bytes === 0) return '0 Bytes'
      const k = 1024
      const sizes = ['Bytes', 'KB', 'MB', 'GB']
      const i = Math.floor(Math.log(bytes) / Math.log(k))
      return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i]
    }

    // Refresh results from backend
    const refreshResults = async () => {
      console.log('[ResultDownload] Refreshing results for job:', props.jobId)
      try {
        const response = await axios.get(`${API_BASE_URL}/api/status/${props.jobId}`)
        if (response.data.results) {
          currentResults.value = response.data.results
          console.log('[ResultDownload] Results refreshed:', currentResults.value)
        }
      } catch (error) {
        console.error('[ResultDownload] Failed to refresh results:', error)
      }
    }

    // Copy MathJax script to clipboard
    const copyMathjaxScript = async () => {
      try {
        await navigator.clipboard.writeText(mathjaxScript)
        copySuccess.value = true
        setTimeout(() => {
          copySuccess.value = false
        }, 2000)
      } catch (error) {
        console.error('[ResultDownload] Failed to copy to clipboard:', error)
        // Fallback
        if (mathjaxTextarea.value) {
          mathjaxTextarea.value.select()
          document.execCommand('copy')
          copySuccess.value = true
          setTimeout(() => {
            copySuccess.value = false
          }, 2000)
        }
      }
    }

    // Download with refresh check
    const downloadFileWithRefresh = async (index) => {
      try {
        // First check if job still exists
        const statusResponse = await axios.get(`${API_BASE_URL}/api/status/${props.jobId}`)

        if (statusResponse.data.error) {
          alert('Job has expired. Please process your document again.')
          return
        }

        // Update results from status response
        if (statusResponse.data.results) {
          currentResults.value = statusResponse.data.results
        }

        // Now download
        const url = `${API_BASE_URL}/api/download/${props.jobId}/${index}`
        console.log('[ResultDownload] Downloading file from:', url)

        const link = document.createElement('a')
        link.href = url
        link.download = currentResults.value[index]?.filename || `download_${index}`
        document.body.appendChild(link)
        link.click()
        document.body.removeChild(link)
      } catch (error) {
        console.error('[ResultDownload] Download failed:', error)
        if (error.response?.status === 404) {
          jobExpired.value = true
          alert('Job has expired or was not found. Click "Process More Documents" to start over.')
        } else {
          alert('Download failed. Please try again.')
        }
      }
    }

    const downloadAll = async () => {
      try {
        // First check if job still exists
        const statusResponse = await axios.get(`${API_BASE_URL}/api/status/${props.jobId}`)

        if (statusResponse.data.error) {
          alert('Job has expired. Please process your document again.')
          return
        }

        const url = `${API_BASE_URL}/api/download/${props.jobId}`
        console.log('[ResultDownload] Downloading all as ZIP from:', url)

        const link = document.createElement('a')
        link.href = url
        link.download = `results_${props.jobId}.zip`
        document.body.appendChild(link)
        link.click()
        document.body.removeChild(link)
      } catch (error) {
        console.error('[ResultDownload] Download all failed:', error)
        if (error.response?.status === 404) {
          jobExpired.value = true
          alert('Job has expired or was not found. Click "Process More Documents" to start over.')
        } else {
          alert('Download failed. Please try again.')
        }
      }
    }

    // Auto-refresh on mount
    onMounted(() => {
      console.log('[ResultDownload] Component mounted, auto-refreshing results...')
      setTimeout(refreshResults, 1000)
    })

    return {
      downloadFileWithRefresh,
      downloadAll,
      hasErrors,
      formatFileSize,
      refreshResults,
      currentResults,
      mathjaxScript,
      mathjaxTextarea,
      copySuccess,
      copyMathjaxScript,
      jobExpired
    }
  }
}
</script>
