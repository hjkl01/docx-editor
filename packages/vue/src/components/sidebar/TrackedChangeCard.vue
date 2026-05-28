<!--
  Mirror of packages/react/src/components/sidebar/TrackedChangeCard.tsx.
  Same chrome (collapsed/expanded), same author/date layout,
  insertion/deletion/replacement formatting, accept/reject icon
  buttons in expanded state.
-->
<template>
  <div
    class="tc-card"
    :class="{ 'tc-card--expanded': expanded }"
    @click="$emit('click')"
    @mousedown.stop
  >
    <div class="tc-card__head">
      <Avatar :name="authorName" :size="32" />
      <div class="tc-card__author-block">
        <div class="tc-card__author">{{ authorName }}</div>
        <div v-if="change.date" class="tc-card__date">{{ formatDate(change.date) }}</div>
      </div>
      <div v-if="expanded" class="tc-card__actions">
        <button class="tc-card__icon-btn" title="Accept" @click.stop="onAccept">
          <MaterialSymbol name="check" :size="20" />
        </button>
        <button class="tc-card__icon-btn" title="Reject" @click.stop="onReject">
          <MaterialSymbol name="close" :size="20" />
        </button>
      </div>
    </div>

    <div class="tc-card__body">
      <template v-if="change.type === 'replacement'">
        Replaced
        <span class="tc-card__deleted">&quot;{{ truncateText(change.deletedText || '') }}&quot;</span>
        with
        <span class="tc-card__inserted">&quot;{{ truncateText(change.text) }}&quot;</span>
      </template>
      <template v-else-if="change.type === 'paragraphMarkInsertion'">
        Inserted paragraph break<template v-if="change.text">: <span class="tc-card__inserted">&quot;{{ truncateText(change.text) }}&quot;</span></template>
      </template>
      <template v-else-if="change.type === 'paragraphMarkDeletion'">
        Deleted paragraph break<template v-if="change.text">: <span class="tc-card__deleted">&quot;{{ truncateText(change.text) }}&quot;</span></template>
      </template>
      <template v-else-if="change.type === 'paragraphPropertiesChanged'">
        Changed paragraph properties<template v-if="change.text">: <span class="tc-card__changed">&quot;{{ truncateText(change.text) }}&quot;</span></template>
      </template>
      <template v-else-if="change.type === 'rowInserted'">
        <span class="tc-card__inserted">Inserted row</span>
      </template>
      <template v-else-if="change.type === 'rowDeleted'">
        <span class="tc-card__deleted">Deleted row</span>
      </template>
      <template v-else-if="change.type === 'cellInserted'">
        <span class="tc-card__inserted">Inserted cell</span>
      </template>
      <template v-else-if="change.type === 'cellDeleted'">
        <span class="tc-card__deleted">Deleted cell</span>
      </template>
      <template v-else-if="change.type === 'cellMerged'">
        <span class="tc-card__changed">Merged cells</span>
      </template>
      <template v-else-if="change.type === 'rowPropertiesChanged'">
        <span class="tc-card__changed">Changed row properties</span>
      </template>
      <template v-else-if="change.type === 'cellPropertiesChanged'">
        <span class="tc-card__changed">Changed cell properties</span>
      </template>
      <template v-else-if="change.type === 'tablePropertiesChanged'">
        <span class="tc-card__changed">Changed table properties</span>
      </template>
      <template v-else>
        {{ change.type === 'insertion' ? 'Added' : 'Deleted' }}
        <span :class="change.type === 'insertion' ? 'tc-card__inserted' : 'tc-card__deleted'">
          &quot;{{ truncateText(change.text) }}&quot;
        </span>
      </template>
    </div>

    <!-- Reply input — mirrors React TrackedChangeCard.tsx:103. Lets
         a user thread a comment under a tracked change. -->
    <ReplyInput
      v-if="expanded"
      @submit="(text: string) => $emit('reply', change.revisionId, text)"
    />
  </div>
</template>

<script setup lang="ts">
import { computed } from 'vue';
import type { TrackedChangeEntry } from './sidebarUtils';
import { formatDate, truncateText } from './sidebarUtils';
import Avatar from './Avatar.vue';
import MaterialSymbol from '../ui/MaterialSymbol.vue';
import ReplyInput from './ReplyInput.vue';

const props = defineProps<{
  change: TrackedChangeEntry;
  expanded: boolean;
}>();

const emit = defineEmits<{
  (e: 'click'): void;
  (e: 'accept', from: number, to: number): void;
  (e: 'reject', from: number, to: number): void;
  (e: 'accept-by-id', revisionId: number): void;
  (e: 'reject-by-id', revisionId: number): void;
  (e: 'reply', revisionId: number, text: string): void;
}>();

const authorName = computed(() => props.change.author || 'Unknown');

const isParagraphMark = computed(() => {
  const t = props.change.type;
  return (
    t === 'paragraphMarkInsertion' ||
    t === 'paragraphMarkDeletion' ||
    t === 'paragraphPropertiesChanged' ||
    t === 'rowInserted' ||
    t === 'rowDeleted' ||
    t === 'rowPropertiesChanged' ||
    t === 'cellInserted' ||
    t === 'cellDeleted' ||
    t === 'cellMerged' ||
    t === 'cellPropertiesChanged' ||
    t === 'tablePropertiesChanged'
  );
});

function onAccept() {
  if (isParagraphMark.value) emit('accept-by-id', props.change.revisionId);
  else emit('accept', props.change.from, props.change.to);
}

function onReject() {
  if (isParagraphMark.value) emit('reject-by-id', props.change.revisionId);
  else emit('reject', props.change.from, props.change.to);
}
</script>

<style scoped>
.tc-card {
  padding: 8px 10px;
  border-radius: 8px;
  background: #f8fbff;
  cursor: pointer;
  box-shadow:
    0 1px 3px rgba(60, 64, 67, 0.2),
    0 2px 6px rgba(60, 64, 67, 0.08);
  margin-bottom: 6px;
  transition: box-shadow 0.15s ease, background-color 0.15s ease, padding 0.15s ease;
}
.tc-card--expanded {
  padding: 10px 12px;
  background: #fff;
  box-shadow:
    0 1px 3px rgba(60, 64, 67, 0.3),
    0 4px 8px 3px rgba(60, 64, 67, 0.15);
}
.tc-card__head {
  display: flex;
  align-items: flex-start;
  gap: 10px;
}
.tc-card__author-block {
  flex: 1;
  min-width: 0;
}
.tc-card__author {
  font-size: 13px;
  font-weight: 600;
  color: #202124;
}
.tc-card__date {
  font-size: 11px;
  color: #5f6368;
}
.tc-card__actions {
  display: flex;
  gap: 4px;
  margin-top: 2px;
}
.tc-card__icon-btn {
  background: none;
  border: none;
  cursor: pointer;
  padding: 4px;
  color: #5f6368;
  display: flex;
  border-radius: 50%;
}
.tc-card__icon-btn:hover {
  background: rgba(60, 64, 67, 0.08);
}
.tc-card__body {
  font-size: 13px;
  line-height: 20px;
  color: #202124;
  margin-top: 6px;
}
.tc-card__deleted {
  color: #c5221f;
  font-weight: 500;
}
.tc-card__inserted {
  color: #137333;
  font-weight: 500;
}
.tc-card__changed {
  color: #5f6368;
  font-weight: 500;
}
</style>
