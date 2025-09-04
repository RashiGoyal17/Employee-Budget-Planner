import { ComponentFixture, TestBed } from '@angular/core/testing';

import { Spreadsheet } from './spreadsheet';

describe('Spreadsheet', () => {
  let component: Spreadsheet;
  let fixture: ComponentFixture<Spreadsheet>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [Spreadsheet]
    })
    .compileComponents();

    fixture = TestBed.createComponent(Spreadsheet);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
