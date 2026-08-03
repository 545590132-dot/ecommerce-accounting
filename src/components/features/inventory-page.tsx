'use client';

import React, { useState, useMemo, useCallback, useRef } from 'react';
import { useAppStore } from '@/store';
import type { InventoryRecord } from '@/types';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Badge } from '@/components/ui/badge';
import { Popover, PopoverContent, PopoverTrigger } from '@/components/ui/popover';
import { Checkbox } from '@/components/ui/checkbox';
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select';
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogFooter,
} from '@/components/ui/dialog';
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from '@/components/ui/table';
import {
  Upload,
  Download,
  Trash2,
  ChevronUp,
  ChevronDown,
  Filter,
  X,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

interface DisplayRow {
  sku: string;
  productName: string;
  productOwner: string;
  stock: number;
  monthlySales: number;
  salesStatus: InventoryRecord['salesStatus'];
  displaySalesStatus: '热销' | '正常' | '平销' | '清货' | '';
  estimatedMonths: number | null;
  goodsValue: number;
  recordIds: string[];
  yearMonth: string;
  adjustmentPlan: string;
}

function formatNumber(n: number, decimals = 0): string {
  if (n === 0) return '0';
  const abs = Math.abs(n);
  if (abs >= 10000) {
    return n.toLocaleString('zh-CN', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
  }
  return n.toLocaleString('zh-CN', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

function formatCurrency(n: number): string {
  return n.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export default function InventoryPage() {
  const {
    skuMappings,
    inventoryFiles,
    inventoryRecords,
    importInventory,
    deleteInventoryFile,
    batchUpdateInventorySalesStatus,
    updateInventoryAdjustmentPlan,
    calculateSummary,
  } = useAppStore();

  // ====== 导入弹窗 ======
  const [importOpen, setImportOpen] = useState(false);
  const [importYear, setImportYear] = useState('');
  const [importMonth, setImportMonth] = useState('');
  const [importFileName, setImportFileName] = useState('');
  const [importData, setImportData] = useState<{ sku: string; stockQty: number }[]>([]);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // ====== 筛选 ======
  const [yearMonthFilter, setYearMonthFilter] = useState<string[]>([]);
  const [ownerFilter, setOwnerFilter] = useState<string[]>([]);
  const [statusFilter, setStatusFilter] = useState<string[]>([]);
  const [ymPopoverOpen, setYmPopoverOpen] = useState(false);
  const [ownerPopoverOpen, setOwnerPopoverOpen] = useState(false);
  const [statusPopoverOpen, setStatusPopoverOpen] = useState(false);
  const [hideZeroStock, setHideZeroStock] = useState(false);

  // ====== 排序 ======
  type SortKey = 'skuCode' | 'stock' | 'monthlySales' | 'estimatedMonths' | 'goodsValue';

  // 从SKU编码前4位提取排序键：数字优先升序，纯英文排最后
  const getSkuSortKey = (sku: string): { num: number; prefix: string } => {
    const prefix = sku.slice(0, 4);
    const digits = prefix.match(/\d+/);
    if (digits) {
      return { num: parseInt(digits[0], 10), prefix };
    }
    return { num: Infinity, prefix };
  };

  const getSkuDisplayCode = (sku: string): string => {
    const prefix = sku.slice(0, 4);
    const digits = prefix.match(/\d+/);
    if (digits) return digits[0];
    return prefix;
  };
  type SortDir = 'asc' | 'desc';
  const [sortKey, setSortKey] = useState<SortKey | null>(null);
  const [sortDir, setSortDir] = useState<SortDir>('desc');

  // ====== 调整计划编辑状态 ======
  const [editingPlanKey, setEditingPlanKey] = useState<string | null>(null);
  const [editingPlanValue, setEditingPlanValue] = useState('');

  const handleSort = (key: SortKey) => {
    if (sortKey === key) {
      setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortKey(key);
      setSortDir('desc');
    }
  };

  const handleSaveAdjustmentPlan = async (recordIds: string[], value: string) => {
    await updateInventoryAdjustmentPlan(recordIds, value);
    setEditingPlanKey(null);
    setEditingPlanValue('');
  };

  const SortIcon = ({ column }: { column: SortKey }) => {
    if (sortKey !== column) return <ChevronUp className="w-3 h-3 text-slate-300" />;
    return sortDir === 'asc' ? (
      <ChevronUp className="w-3 h-3 text-slate-800" />
    ) : (
      <ChevronDown className="w-3 h-3 text-slate-800" />
    );
  };

  // ====== 可用筛选选项 ======
  const availableYearMonths = useMemo(() => {
    const set = new Set<string>();
    for (const f of inventoryFiles) set.add(f.yearMonth);
    return Array.from(set).sort().reverse();
  }, [inventoryFiles]);

  const availableOwners = useMemo(() => {
    const skuOwnerMap = new Map<string, string>();
    for (const m of skuMappings) {
      const key = m.sku.toLowerCase().replace(/\s/g, '');
      if (m.productOwner) skuOwnerMap.set(key, m.productOwner);
    }
    const owners = new Set<string>();
    for (const r of inventoryRecords) {
      const key = r.sku.toLowerCase().replace(/\s/g, '');
      const owner = skuOwnerMap.get(key);
      if (owner) owners.add(owner);
    }
    return Array.from(owners).sort();
  }, [skuMappings, inventoryRecords]);

  const availableStatuses: string[] = ['清货', '热销', '平销', '正常']; // 筛选包含所有可能的最终状态

  // ====== 月销量查找：从三平台订单数据中汇总 ======
  const salesQtyMap = useMemo(() => {
    const map = new Map<string, number>(); // key: `${yearMonth}__${normalizedSku}` -> totalQuantity
    const platforms: ('shopee' | 'lazada' | 'tiktok')[] = ['shopee', 'lazada', 'tiktok'];
    for (const p of platforms) {
      const summary = calculateSummary(p);
      for (const order of summary.orders) {
        const key = `${order.orderDate}__${order.sku.toLowerCase().replace(/\s/g, '')}`;
        map.set(key, (map.get(key) || 0) + order.quantity);
      }
    }
    return map;
  }, [calculateSummary]);

  // ====== 构建展示数据 ======
  const displayRows = useMemo(() => {
    const skuMap = new Map<string, { productName: string; productOwner: string; purchasePrice: number }>();
    for (const m of skuMappings) {
      const key = m.sku.toLowerCase().replace(/\s/g, '');
      skuMap.set(key, {
        productName: m.productName,
        productOwner: m.productOwner || '',
        purchasePrice: m.purchasePrice,
      });
    }

    // DEBUG: 分组前检查重复SKU
    const skuCountByMonth = new Map<string, Map<string, number>>();
    for (const r of inventoryRecords) {
      const monthMap = skuCountByMonth.get(r.yearMonth) || new Map();
      monthMap.set(r.sku, (monthMap.get(r.sku) || 0) + 1);
      skuCountByMonth.set(r.yearMonth, monthMap);
    }
    for (const [ym, monthMap] of skuCountByMonth) {
      const dupes: [string, number][] = [];
      for (const [sku, cnt] of monthMap) {
        if (cnt > 1) dupes.push([sku, cnt]);
      }
      if (dupes.length > 0) {
        console.log(`[重复SKU] ${ym}: ${dupes.length}个SKU有重复`, dupes.slice(0, 10).map(d => `SKU="${d[0]}"(${d[1]}次)`));
      } else {
        console.log(`[重复SKU] ${ym}: 无重复, 共${monthMap.size}个SKU`);
      }
    }

    // 按 yearMonth + sku 分组，累加同一SKU的库存（处理重复记录）
    const grouped = new Map<string, { sku: string; stockQty: number; yearMonth: string; salesStatus: InventoryRecord['salesStatus']; recordIds: string[]; adjustmentPlan: string }>();
    for (const r of inventoryRecords) {
      const groupKey = `${r.yearMonth}__${r.sku}`;
      const existing = grouped.get(groupKey);
      if (!existing) {
        grouped.set(groupKey, {
          sku: r.sku,
          stockQty: r.stockQty,
          yearMonth: r.yearMonth,
          salesStatus: r.salesStatus,
          recordIds: [r.id],
          adjustmentPlan: r.adjustmentPlan || '',
        });
      } else {
        // 同一(yearMonth, sku)有多条记录时，累加库存
        existing.stockQty += r.stockQty;
        existing.recordIds.push(r.id);
        // 保留最新的销售情况（非默认值优先）
        if (r.salesStatus) {
          existing.salesStatus = r.salesStatus;
        }
        if (r.adjustmentPlan) {
          existing.adjustmentPlan = r.adjustmentPlan;
        }
      }
    }

    // DEBUG: 分组后校验
    console.log(`[显示] inventoryRecords=${inventoryRecords.length}, grouped=${grouped.size}`);
    const groupedMonthTotals = new Map<string, { count: number; total: number }>();
    for (const [, v] of grouped) {
      const mt = groupedMonthTotals.get(v.yearMonth) || { count: 0, total: 0 };
      mt.count++;
      mt.total += v.stockQty;
      groupedMonthTotals.set(v.yearMonth, mt);
    }
    for (const [ym, mt] of groupedMonthTotals) {
      console.log(`[显示] grouped ${ym}: ${mt.count}组, 总库存=${mt.total}`);
    }

    const rows: DisplayRow[] = [];
    for (const [, v] of grouped) {
      const skuKey = v.sku.toLowerCase().replace(/\s/g, '');
      const mapping = skuMap.get(skuKey);
      const productName = mapping?.productName || v.sku;
      const productOwner = mapping?.productOwner || '';

      // 月销量 = 三平台订单中该月份该SKU的销量求和
      const salesKey = `${v.yearMonth}__${skuKey}`;
      const monthlySales = salesQtyMap.get(salesKey) || 0;

      // 预估销售时长 = 库存 / 月销量（月销量>0时）
      const estimatedMonths = monthlySales > 0 ? v.stockQty / monthlySales : null;

      // 计算 displaySalesStatus
      const displaySalesStatus: '清货' | '热销' | '正常' | '平销' = (v.salesStatus === '清货' ? '清货' : (monthlySales >= 500 ? '热销' : (estimatedMonths !== null && estimatedMonths > 0 && estimatedMonths <= 6 ? '正常' : '平销'))) as '清货' | '热销' | '正常' | '平销';

      // 筛选（月份、负责人、零库存）- 销售情况筛选移到合并后
      if (yearMonthFilter.length > 0 && !yearMonthFilter.includes(v.yearMonth)) continue;
      if (ownerFilter.length > 0 && !ownerFilter.includes(productOwner)) continue;
      if (hideZeroStock && v.stockQty === 0) continue;

      // 货值 = 库存 * 采购成本
      const goodsValue = v.stockQty * (mapping?.purchasePrice || 0);

      rows.push({
        sku: v.sku,
        productName,
        productOwner,
        stock: v.stockQty,
        monthlySales,
        salesStatus: v.salesStatus,
        displaySalesStatus,
        estimatedMonths,
        goodsValue,
        recordIds: v.recordIds,
        yearMonth: v.yearMonth,
        adjustmentPlan: v.adjustmentPlan || '',
      });
    }

    // 合并相同商品名称的行
    const mergedMap = new Map<string, {
      sku: string;
      productName: string;
      productOwner: string;
      stock: number;
      monthlySales: number;
      salesStatus: InventoryRecord['salesStatus'];
      goodsValue: number;
      yearMonth: string;
      recordIds: string[];
      adjustmentPlan: string;
    }>();

    for (const row of rows) {
      const mergeKey = `${row.yearMonth}__${row.productName}`;
      const existing = mergedMap.get(mergeKey);
      if (existing) {
        existing.stock += row.stock;
        existing.monthlySales += row.monthlySales;
        existing.goodsValue += row.goodsValue;
        // 保留最后一个销售状态
        if (row.salesStatus) existing.salesStatus = row.salesStatus;
        existing.recordIds.push(...row.recordIds);
        // 保留最后一个非空调整计划
        if (row.adjustmentPlan) existing.adjustmentPlan = row.adjustmentPlan;
      } else {
        mergedMap.set(mergeKey, {
          sku: row.sku,
          productName: row.productName,
          productOwner: row.productOwner,
          stock: row.stock,
          monthlySales: row.monthlySales,
          salesStatus: row.salesStatus,
          goodsValue: row.goodsValue,
          yearMonth: row.yearMonth,
          recordIds: row.recordIds,
          adjustmentPlan: row.adjustmentPlan,
        });
      }
    }

    // 转换为最终展示行，计算系统判定
    const mergedRows: DisplayRow[] = [];
    // [DEBUG] 合并后统计
    const mergedMonthlyTotals = new Map<string, { count: number; total: number }>();
    for (const [, v] of mergedMap) {
      const ym = v.yearMonth;
      const mt = mergedMonthlyTotals.get(ym) || { count: 0, total: 0 };
      mt.count++;
      mt.total += v.stock;
      mergedMonthlyTotals.set(ym, mt);
    }
    console.log(`[显示管道] 合并后: ${mergedMap.size}行`);
    for (const [ym, mt] of mergedMonthlyTotals) {
      console.log(`[显示管道] 合并后 ${ym}: ${mt.count}行, 总库存=${mt.total}`);
    }
    for (const [, v] of mergedMap) {
      const estimatedMonths = v.monthlySales > 0 ? v.stock / v.monthlySales : null;

      // 计算实际显示的销售情况
      let displaySalesStatus: '热销' | '正常' | '平销' | '清货' | '' = '';
      if (v.salesStatus === '清货') {
        displaySalesStatus = '清货';
      } else if (v.salesStatus === '系统判定') {
        if (v.monthlySales >= 500) {
          displaySalesStatus = '热销';
        } else if (estimatedMonths !== null && estimatedMonths <= 6) {
          displaySalesStatus = '正常';
        } else {
          displaySalesStatus = '平销';
        }
      }

      mergedRows.push({
        sku: v.sku,
        productName: v.productName,
        productOwner: v.productOwner,
        stock: v.stock,
        monthlySales: v.monthlySales,
        salesStatus: v.salesStatus,
        displaySalesStatus,
        estimatedMonths,
        goodsValue: v.goodsValue,
        recordIds: v.recordIds,
        yearMonth: v.yearMonth,
        adjustmentPlan: v.adjustmentPlan,
      });
    }

    // 隐藏库存为0的商品 + 销售情况筛选（合并后）
    let filtered = mergedRows;
    if (statusFilter.length > 0) {
      filtered = filtered.filter(r => r.displaySalesStatus && statusFilter.includes(r.displaySalesStatus));
    }
    if (hideZeroStock) {
      filtered = filtered.filter(r => r.stock > 0);
    }

    // DEBUG: 合并和筛选后校验
    const mergedTotal = mergedRows.reduce((s, r) => s + r.stock, 0);
    const filteredTotal = filtered.reduce((s, r) => s + r.stock, 0);
    console.log(`[显示] mergedRows=${mergedRows.length}, total=${mergedTotal}, filtered=${filtered.length}, filteredTotal=${filteredTotal}, hideZeroStock=${hideZeroStock}, ymFilter=[${yearMonthFilter}]`);

    // 排序
    if (sortKey) {
      filtered.sort((a, b) => {
        if (sortKey === 'skuCode') {
          const ka = getSkuSortKey(a.sku);
          const kb = getSkuSortKey(b.sku);
          const cmp = ka.num - kb.num || ka.prefix.localeCompare(kb.prefix);
          return sortDir === 'asc' ? cmp : -cmp;
        }
        const av = a[sortKey] ?? -Infinity;
        const bv = b[sortKey] ?? -Infinity;
        return sortDir === 'asc' ? (av as number) - (bv as number) : (bv as number) - (av as number);
      });
    }

    // [DEBUG] 筛选后统计
    const filteredTotals = new Map<string, { count: number; total: number }>();
    for (const r of filtered) {
      const ym = r.yearMonth;
      const ft = filteredTotals.get(ym) || { count: 0, total: 0 };
      ft.count++;
      ft.total += r.stock;
      filteredTotals.set(ym, ft);
    }
    console.log(`[显示管道] 筛选后: ${filtered.length}行`);
    for (const [ym, ft] of filteredTotals) {
      console.log(`[显示管道] 筛选后 ${ym}: ${ft.count}行, 总库存=${ft.total}`);
    }

    return filtered;
  }, [skuMappings, inventoryRecords, yearMonthFilter, ownerFilter, statusFilter, sortKey, sortDir, hideZeroStock]);

  // ====== 总计行 ======
  const totalRow = useMemo(() => {
    const totalStock = displayRows.reduce((s, r) => s + r.stock, 0);
    const totalMonthlySales = displayRows.reduce((s, r) => s + r.monthlySales, 0);
    const totalGoodsValue = displayRows.reduce((s, r) => s + r.goodsValue, 0);
    const totalEstimatedMonths = totalMonthlySales > 0 ? totalStock / totalMonthlySales : null;
    // DEBUG: 直接从inventoryRecords计算总库存（不经过显示管道），用于对比
    const directTotal = inventoryRecords
      .filter(r => yearMonthFilter.length === 0 || yearMonthFilter.includes(r.yearMonth))
      .reduce((s, r) => s + r.stockQty, 0);
    if (directTotal !== totalStock) {
      console.error(`[数据不一致!] 直接计算=${directTotal}, 管道计算=${totalStock}, 差异=${directTotal - totalStock}`);
    } else {
      console.log(`[数据一致] 总库存=${totalStock}`);
    }
    return { totalStock, totalMonthlySales, totalGoodsValue, totalEstimatedMonths };
  }, [displayRows]);

  // ====== 导入处理 ======
  const handleFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target?.result as ArrayBuffer);
      const wb = XLSX.read(data, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws);

      const records: { sku: string; stockQty: number }[] = [];
      for (const row of json) {
        // 查找"商品名称"或"SKU编码"列
        const skuKey = Object.keys(row).find(
          (k) => k.trim() === '商品名称' || k.trim() === 'SKU编码' || k.trim() === 'SKU' || k.trim().toLowerCase() === 'sku'
        );
        // 查找"可用量"或"库存"列
        const stockKey = Object.keys(row).find(
          (k) => k.trim() === '可用量' || k.trim() === '库存量' || k.trim() === '库存' || k.trim() === '可用库存'
        );
        if (skuKey && stockKey) {
          const sku = String(row[skuKey] ?? '').trim();
          const stockQty = Number(row[stockKey]) || 0;
          if (sku) records.push({ sku, stockQty });
        }
      }
      setImportData(records);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleImportConfirm = useCallback(async () => {
    if (!importYear || !importMonth || importData.length === 0) return;
    const yearMonth = `${importYear}-${importMonth.padStart(2, '0')}`;
    const success = await importInventory({ fileName: importFileName, yearMonth }, importData);
    if (!success) {
      // 保存失败时仍然关闭弹窗（本地状态已更新），但提示用户
      alert('库存数据导入到云端失败，刷新页面后数据可能丢失，请检查网络后重试');
    }
    setImportOpen(false);
    setImportData([]);
    setImportFileName('');
    setImportYear('');
    setImportMonth('');
  }, [importYear, importMonth, importData, importFileName, importInventory]);

  // ====== 导出 ======
  const handleExport = useCallback(() => {
    type ExportRow = Record<string, string | number>;
    const exportData: ExportRow[] = displayRows.map((r, i) => ({
      '序号': i + 1,
      '月份': r.yearMonth,
      '商品编号': getSkuDisplayCode(r.sku),
      '商品名称': r.productName,
      '产品负责人': r.productOwner,
      '月底库存': r.stock,
      '月销量': r.monthlySales,
      '销售情况': r.displaySalesStatus || '',
      '预估销售时长(月)': r.estimatedMonths !== null ? Number(r.estimatedMonths.toFixed(1)) : '',
      '货值': Number(r.goodsValue.toFixed(2)),
      '调整计划': r.adjustmentPlan || '',
    }));
    // 总计行
    exportData.push({
      '序号': '',
      '月份': '',
      '商品编号': '',
      '商品名称': '总计',
      '产品负责人': '所有',
      '月底库存': totalRow.totalStock,
      '月销量': totalRow.totalMonthlySales,
      '销售情况': '',
      '预估销售时长(月)': totalRow.totalEstimatedMonths !== null ? Number(totalRow.totalEstimatedMonths.toFixed(1)) : '',
      '货值': Number(totalRow.totalGoodsValue.toFixed(2)),
      '调整计划': '',
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '库存查询');
    const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    saveAs(new Blob([buf], { type: 'application/octet-stream' }), '库存查询.xlsx');
  }, [displayRows, totalRow]);

  // ====== 模板下载 ======
  const handleDownloadTemplate = useCallback(() => {
    const template = [{ '商品名称': 'SKU001', '可用量': 100 }];
    const ws = XLSX.utils.json_to_sheet(template);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '库存导入模板');
    const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    saveAs(new Blob([buf], { type: 'application/octet-stream' }), '库存导入模板.xlsx');
  }, []);

  // ====== 筛选重置 ======
  const resetFilters = () => {
    setYearMonthFilter([]);
    setOwnerFilter([]);
    setStatusFilter([]);
    setHideZeroStock(false);
  };

  const hasActiveFilter = yearMonthFilter.length > 0 || ownerFilter.length > 0 || statusFilter.length > 0 || hideZeroStock;

  // ====== 年份列表 ======
  const currentYear = new Date().getFullYear();
  const years = Array.from({ length: 5 }, (_, i) => String(currentYear - 2 + i));
  const months = Array.from({ length: 12 }, (_, i) => String(i + 1).padStart(2, '0'));

  return (
    <div className="space-y-4">
      {/* 操作栏 */}
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div className="flex items-center gap-2">
          <Button variant="outline" size="sm" onClick={() => setImportOpen(true)}>
            <Upload className="w-4 h-4 mr-1" /> 导入库存
          </Button>
          <Button variant="outline" size="sm" onClick={handleExport} disabled={displayRows.length === 0}>
            <Download className="w-4 h-4 mr-1" /> 导出
          </Button>
          <Button variant="outline" size="sm" onClick={handleDownloadTemplate}>
            下载模板
          </Button>
        </div>
        <div className="flex items-center gap-2 text-sm text-slate-500">
          已导入 {inventoryFiles.length} 个库存文件
        </div>
      </div>

      {/* 已导入文件列表 */}
      {inventoryFiles.length > 0 && (
        <div className="flex flex-wrap gap-2">
          {inventoryFiles.map((f) => (
            <div key={f.id} className="flex items-center gap-1 px-2 py-1 bg-slate-50 rounded-md border text-sm">
              <span className="text-slate-700">{f.fileName}</span>
              <Badge variant="outline" className="text-xs">{f.yearMonth}</Badge>
              <button
                className="ml-1 text-slate-400 hover:text-red-500 transition-colors"
                onClick={() => deleteInventoryFile(f.id)}
              >
                <Trash2 className="w-3 h-3" />
              </button>
            </div>
          ))}
        </div>
      )}

      {/* 筛选栏 */}
      <div className="flex items-center gap-2 flex-wrap">
        {/* 年月筛选 */}
        <Popover open={ymPopoverOpen} onOpenChange={setYmPopoverOpen}>
          <PopoverTrigger asChild>
            <Button variant="outline" size="sm" className="h-8">
              <Filter className="w-3 h-3 mr-1" />
              {yearMonthFilter.length === 0 ? '月销月份' : `月份(${yearMonthFilter.length})`}
            </Button>
          </PopoverTrigger>
          <PopoverContent className="w-48 p-2" align="start">
            <div className="space-y-1 max-h-60 overflow-y-auto">
              {availableYearMonths.length === 0 && (
                <div className="text-sm text-slate-400 px-2 py-1">暂无数据</div>
              )}
              {availableYearMonths.map((ym) => (
                <label key={ym} className="flex items-center gap-2 px-2 py-1 hover:bg-slate-50 rounded cursor-pointer text-sm">
                  <Checkbox
                    checked={yearMonthFilter.includes(ym)}
                    onCheckedChange={(checked) => {
                      setYearMonthFilter((prev) =>
                        checked ? [...prev, ym] : prev.filter((v) => v !== ym)
                      );
                    }}
                  />
                  {ym}
                </label>
              ))}
            </div>
          </PopoverContent>
        </Popover>

        {/* 产品负责人筛选 */}
        {availableOwners.length > 0 && (
          <Popover open={ownerPopoverOpen} onOpenChange={setOwnerPopoverOpen}>
            <PopoverTrigger asChild>
              <Button variant="outline" size="sm" className="h-8">
                <Filter className="w-3 h-3 mr-1" />
                {ownerFilter.length === 0 ? '全部负责人' : `负责人(${ownerFilter.length})`}
              </Button>
            </PopoverTrigger>
            <PopoverContent className="w-48 p-2" align="start">
              <div className="space-y-1 max-h-60 overflow-y-auto">
                {availableOwners.map((owner) => (
                  <label key={owner} className="flex items-center gap-2 px-2 py-1 hover:bg-slate-50 rounded cursor-pointer text-sm">
                    <Checkbox
                      checked={ownerFilter.includes(owner)}
                      onCheckedChange={(checked) => {
                        setOwnerFilter((prev) =>
                          checked ? [...prev, owner] : prev.filter((v) => v !== owner)
                        );
                      }}
                    />
                    {owner}
                  </label>
                ))}
              </div>
            </PopoverContent>
          </Popover>
        )}

        {/* 销售情况筛选 */}
        <Popover open={statusPopoverOpen} onOpenChange={setStatusPopoverOpen}>
          <PopoverTrigger asChild>
            <Button variant="outline" size="sm" className="h-8">
              <Filter className="w-3 h-3 mr-1" />
              {statusFilter.length === 0 ? '全部销售情况' : `销售情况(${statusFilter.length})`}
            </Button>
          </PopoverTrigger>
          <PopoverContent className="w-48 p-2" align="start">
            <div className="space-y-1">
              {availableStatuses.map((status) => (
                <label key={status} className="flex items-center gap-2 px-2 py-1 hover:bg-slate-50 rounded cursor-pointer text-sm">
                  <Checkbox
                    checked={statusFilter.includes(status)}
                    onCheckedChange={(checked) => {
                      setStatusFilter((prev) =>
                        checked ? [...prev, status] : prev.filter((v) => v !== status)
                      );
                    }}
                  />
                  {status}
                </label>
              ))}
            </div>
          </PopoverContent>
        </Popover>

        <label className="flex items-center gap-1.5 text-sm text-slate-600 cursor-pointer">
          <Checkbox
            checked={hideZeroStock}
            onCheckedChange={(checked) => setHideZeroStock(checked === true)}
          />
          隐藏零库存
        </label>

        {hasActiveFilter && (
          <Button variant="ghost" size="sm" className="h-8 text-slate-500" onClick={resetFilters}>
            <X className="w-3 h-3 mr-1" /> 清除筛选
          </Button>
        )}
      </div>

      {/* 数据表格 */}
      <div className="border rounded-lg overflow-auto">
        <Table>
          <TableHeader>
            <TableRow className="bg-slate-50">
              <TableHead className="text-center w-[50px]">序号</TableHead>
              <TableHead className="text-center w-[80px]">月份</TableHead>
              <TableHead
                className="cursor-pointer select-none text-center"
                onClick={() => handleSort('skuCode')}
              >
                <span className="inline-flex items-center gap-1">
                  商品编号 <SortIcon column="skuCode" />
                </span>
              </TableHead>
              <TableHead>商品名称</TableHead>
              <TableHead>产品负责人</TableHead>
              <TableHead
                className="cursor-pointer select-none text-right"
                onClick={() => handleSort('stock')}
              >
                <span className="inline-flex items-center gap-1">
                  月底库存 <SortIcon column="stock" />
                </span>
              </TableHead>
              <TableHead
                className="cursor-pointer select-none text-right"
                onClick={() => handleSort('monthlySales')}
              >
                <span className="inline-flex items-center gap-1">
                  月销量 <SortIcon column="monthlySales" />
                </span>
              </TableHead>
              <TableHead className="text-center">销售情况</TableHead>
              <TableHead
                className="cursor-pointer select-none text-right"
                onClick={() => handleSort('estimatedMonths')}
              >
                <span className="inline-flex items-center gap-1">
                  预估销售时长(月) <SortIcon column="estimatedMonths" />
                </span>
              </TableHead>
              <TableHead
                className="cursor-pointer select-none text-right"
                onClick={() => handleSort('goodsValue')}
              >
                <span className="inline-flex items-center gap-1">
                  货值 <SortIcon column="goodsValue" />
                </span>
              </TableHead>
              <TableHead className="text-center min-w-[120px]">调整计划</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {displayRows.length === 0 ? (
              <TableRow>
                <TableCell colSpan={11} className="h-24 text-center text-slate-400">
                  暂无库存数据，请先导入库存表格
                </TableCell>
              </TableRow>
            ) : (
              <>
                {displayRows.map((row, idx) => (
                  <TableRow key={row.recordIds.join(',')}>
                    <TableCell className="text-center text-slate-400 text-sm">{idx + 1}</TableCell>
                    <TableCell className="text-center text-slate-500 text-sm">{row.yearMonth}</TableCell>
                    <TableCell className="text-center text-slate-500 text-sm font-mono">{getSkuDisplayCode(row.sku)}</TableCell>
                    <TableCell className="font-medium">{row.productName}</TableCell>
                    <TableCell className="text-slate-600">{row.productOwner}</TableCell>
                    <TableCell className="text-right font-mono">{formatNumber(row.stock)}</TableCell>
                    <TableCell className="text-right font-mono">
                      <span className={row.monthlySales < 0 ? 'text-red-500' : ''}>
                        {formatNumber(row.monthlySales)}
                      </span>
                    </TableCell>
                    <TableCell className="text-center">
                      <Select
                        value={row.salesStatus || ''}
                        onValueChange={(val) => {
                          const status = val as InventoryRecord['salesStatus'];
                          batchUpdateInventorySalesStatus(row.recordIds, status);
                        }}
                      >
                        <SelectTrigger className="h-7 w-24 mx-auto text-xs">
                          <SelectValue placeholder="选择">
                            {row.salesStatus === '系统判定' && row.displaySalesStatus
                              ? row.displaySalesStatus
                              : row.salesStatus || '选择'}
                          </SelectValue>
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value="清货">清货</SelectItem>
                          <SelectItem value="系统判定">系统判定</SelectItem>
                        </SelectContent>
                      </Select>
                    </TableCell>
                    <TableCell className="text-right font-mono">
                      {row.estimatedMonths !== null ? row.estimatedMonths.toFixed(1) : '-'}
                    </TableCell>
                    <TableCell className="text-right font-mono">{formatCurrency(row.goodsValue)}</TableCell>
                    <TableCell className="text-center">
                      <div className="flex items-center justify-center gap-1">
                        <Input
                          value={editingPlanKey === row.recordIds.join(',') ? editingPlanValue : (row.adjustmentPlan || '')}
                          onChange={(e) => {
                            setEditingPlanKey(row.recordIds.join(','));
                            setEditingPlanValue(e.target.value);
                          }}
                          onFocus={() => {
                            if (editingPlanKey !== row.recordIds.join(',')) {
                              setEditingPlanKey(row.recordIds.join(','));
                              setEditingPlanValue(row.adjustmentPlan || '');
                            }
                          }}
                          className="h-7 text-xs border-slate-200 focus:border-slate-400"
                          placeholder="点击填写"
                        />
                        {editingPlanKey === row.recordIds.join(',') && (
                          <Button
                            size="sm"
                            variant="ghost"
                            className="h-7 px-1.5 text-xs text-slate-500 hover:text-slate-700"
                            onClick={() => {
                              handleSaveAdjustmentPlan(row.recordIds, editingPlanValue);
                              setEditingPlanKey(null);
                              setEditingPlanValue('');
                            }}
                          >
                            保存
                          </Button>
                        )}
                      </div>
                    </TableCell>
                  </TableRow>
                ))}
                {/* 总计行 */}
                <TableRow className="bg-slate-100 font-bold">
                  <TableCell />
                  <TableCell />
                  <TableCell className="text-center text-slate-500">-</TableCell>
                  <TableCell>总计</TableCell>
                  <TableCell>所有</TableCell>
                  <TableCell className="text-right font-mono">{formatNumber(totalRow.totalStock)}</TableCell>
                  <TableCell className="text-right font-mono">{formatNumber(totalRow.totalMonthlySales)}</TableCell>
                  <TableCell />
                  <TableCell className="text-right font-mono">
                    {totalRow.totalEstimatedMonths !== null ? totalRow.totalEstimatedMonths.toFixed(1) : '-'}
                  </TableCell>
                  <TableCell className="text-right font-mono">{formatCurrency(totalRow.totalGoodsValue)}</TableCell>
                  <TableCell />
                </TableRow>
              </>
            )}
          </TableBody>
        </Table>
      </div>

      {/* 导入弹窗 */}
      <Dialog open={importOpen} onOpenChange={setImportOpen}>
        <DialogContent className="sm:max-w-md">
          <DialogHeader>
            <DialogTitle>导入库存数据</DialogTitle>
          </DialogHeader>
          <div className="space-y-4">
            <div>
              <Label>选择文件</Label>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                className="block w-full text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-slate-100 file:text-slate-700 hover:file:bg-slate-200"
                onChange={handleFileSelect}
              />
              <p className="text-xs text-slate-400 mt-1">
                支持 .xlsx/.xls/.csv 格式，需包含「商品名称」和「可用量」列
              </p>
            </div>
            <div className="flex gap-4">
              <div className="flex-1">
                <Label>年份</Label>
                <Select value={importYear} onValueChange={setImportYear}>
                  <SelectTrigger>
                    <SelectValue placeholder="选择年份" />
                  </SelectTrigger>
                  <SelectContent>
                    {years.map((y) => (
                      <SelectItem key={y} value={y}>{y}年</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              <div className="flex-1">
                <Label>月份</Label>
                <Select value={importMonth} onValueChange={setImportMonth}>
                  <SelectTrigger>
                    <SelectValue placeholder="选择月份" />
                  </SelectTrigger>
                  <SelectContent>
                    {months.map((m) => (
                      <SelectItem key={m} value={m}>{parseInt(m)}月</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            </div>
            {importData.length > 0 && (
              <div className="p-3 bg-green-50 rounded-md text-sm text-green-700">
                已解析 {importData.length} 条库存记录
              </div>
            )}
          </div>
          <DialogFooter>
            <Button variant="outline" onClick={() => setImportOpen(false)}>
              取消
            </Button>
            <Button
              onClick={handleImportConfirm}
              disabled={!importYear || !importMonth || importData.length === 0}
            >
              确认导入
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}
