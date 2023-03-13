<template>
  <section class="bg-gray-50 dark:bg-gray-900 p-3 sm:p-5">

    <div v-if="!loading" class="mx-auto max-w-screen-xl px-4 lg:px-12">
      <!-- Start coding here -->
      <div class="bg-white dark:bg-gray-800 relative shadow-md sm:rounded-lg overflow-hidden">
        <div class="flex flex-col md:flex-row items-center justify-between space-y-3 md:space-y-0 md:space-x-4 p-4">
          <div class="w-full md:w-1/2">
            <form class="flex items-center">
              <label for="simple-search" class="sr-only">Search</label>
              <div class="relative w-full">
                <div class="absolute inset-y-0 left-0 flex items-center pl-3 pointer-events-none">
                  <svg aria-hidden="true" class="w-5 h-5 text-gray-500 dark:text-gray-400" fill="currentColor"
                    viewbox="0 0 20 20" xmlns="http://www.w3.org/2000/svg">
                    <path fill-rule="evenodd"
                      d="M8 4a4 4 0 100 8 4 4 0 000-8zM2 8a6 6 0 1110.89 3.476l4.817 4.817a1 1 0 01-1.414 1.414l-4.816-4.816A6 6 0 012 8z"
                      clip-rule="evenodd" />
                  </svg>
                </div>
                <input type="text" id="simple-search" v-model="search"
                  class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-primary-500 focus:border-primary-500 block w-full pl-10 p-2 dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 dark:text-white dark:focus:ring-primary-500 dark:focus:border-primary-500"
                  placeholder="Search" required="">
              </div>
            </form>
          </div>
        </div>
        <div class="overflow-x-auto">
          <table class="w-full text-sm text-left text-gray-500 dark:text-gray-400">
            <thead class="text-xs text-gray-700 uppercase bg-gray-50 dark:bg-gray-700 dark:text-gray-400">
              <tr>
                <th v-for="col in cols" :key="col.key" class="px-4 py-3 text-center">{{ col.name }}</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="item in filteredData" class="border-b dark:border-gray-700">
                <template v-for="(col, indexCol) in cols" :key="col.key">
                  <th v-if="indexCol === 0" scope="row"
                    class="px-4 py-3 font-medium text-center text-gray-900 whitespace-nowrap dark:text-white">{{
                      item[col.key] }}</th>
                  <td v-else class="px-4 py-3 text-center">{{ item[col.key] }}</td>
                </template>
              </tr>

            </tbody>
          </table>
        </div>
      </div>
    </div>
    <div v-if="loading">
      <div class="flex justify-center items-center h-64">
        <div class="loader ease-linear rounded-full border-4 border-t-4 border-gray-200 h-12 w-12 mb-4"></div>
      </div>
    </div>
  </section>
</template>

<script setup>

const { cols, data, loading } = defineProps({
  cols: {
    type: Array,
    required: true
  },
  data: {
    type: Array,
    required: true
  },
  loading: {
    type: Boolean,
    required: true
  }
})

const search = ref('')
const filteredData = computed(() => {
  if (!search.value) return data
  return data.filter(item => {
    return Object.keys(item).some(key => {
      return String(item[key]).toLowerCase().includes(search.value.toLowerCase())
    })
  })
})
</script>

<style lang="scss" scoped></style>
