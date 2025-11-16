import { Component, OnInit } from '@angular/core';
import { Tutorial } from '../../models/tutorial.model';
import { TutorialService } from '../../services/tutorial.service';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-export',
  templateUrl: './export.component.html',
  styleUrls: ['./export.component.css']
})
export class ExportComponent implements OnInit {
  tutorials: Tutorial[] = [];
  sortDirection: 'asc' | 'desc' = 'asc';
  sortColumn: 'id' | 'title' | 'description' | 'published' | null = null;

  constructor(private tutorialService: TutorialService) {}

  ngOnInit(): void {
    this.loadTutorials();
  }

  loadTutorials(): void {
    this.tutorialService.getAll().subscribe({
      next: (data) => {
        this.tutorials = data;
      },
      error: (e) => console.error(e),
    });
  }

  // Generic sort for any column
  sortBy(
    column: 'id' | 'title' | 'description' | 'published',
    type: 'string' | 'number' | 'date' = 'string'
  ): void {
    // Toggle direction when clicking the same column; reset when changing column
    if (this.sortColumn === column) {
      this.sortDirection = this.sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
      this.sortColumn = column;
      this.sortDirection = 'asc';
    }

    this.tutorials = [...this.tutorials].sort((a, b) => {
      let valA: any = (a as any)[column];
      let valB: any = (b as any)[column];

      // Normalize null/undefined
      if (valA === null || valA === undefined) valA = '';
      if (valB === null || valB === undefined) valB = '';

      if (type === 'string') {
        valA = String(valA).toLowerCase();
        valB = String(valB).toLowerCase();
      } else if (type === 'number') {
        valA = Number(valA);
        valB = Number(valB);
      } else if (type === 'date') {
        valA = new Date(valA);
        valB = new Date(valB);
      }

      if (valA < valB) {
        return this.sortDirection === 'asc' ? -1 : 1;
      }
      if (valA > valB) {
        return this.sortDirection === 'asc' ? 1 : -1;
      }
      return 0;
    });
  }

  // Export table data to Excel
  exportToExcel(): void {
    const dataForExcel = this.tutorials.map((t) => ({
      ID: t.id,
      TITLE: t.title,
      DESCRIPTION: t.description,
      PUBLISHED: t.published,
    }));

    const worksheet = XLSX.utils.json_to_sheet(dataForExcel);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Tutorials');

    // Triggers a download (works in all modern browsers)
    XLSX.writeFile(workbook, 'tutorials.xlsx');
  }
}
