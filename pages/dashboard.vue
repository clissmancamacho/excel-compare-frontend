<template>
    <section>
        <CardBox class="mb-6" has-table>
            <FileUpload />
        </CardBox>
        <div v-if="data">
            <CardBox class="mb-6" has-table>
                <TableComponent :data="data" :cols="cols" :loading="loading" />
            </CardBox>
        </div>
    </section>
</template>

<script setup>
import { checkDataStore } from '~~/stores/checkDataStore';
const useCheckDataStore = checkDataStore()
const data = computed(() => useCheckDataStore.data);
const loading = computed(() => useCheckDataStore.loadingData);

const cols = [
    { name: 'Codigo', key: 'code' },
    { name: 'Tipo', key: 'type' },
    { name: 'Monto en Cierre de Caja', key: 'amountCierreCaja' },
    { name: 'Monto en Pago diario', key: 'amountPagoDiario' }
]
setPageLayout('private')
</script>

<style lang="scss" scoped></style>