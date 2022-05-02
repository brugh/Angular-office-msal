import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';

@Component({
  selector: 'app-button',
  templateUrl: './button.component.html',
  styleUrls: ['./button.component.scss']
})
export class ButtonComponent implements OnInit {
  @Input() icon: string = '';
  @Input() label: string = '';
  @Input() appearance: 'accent' | 'lightweight' | 'neutral' | 'outline' |'stealth' = 'neutral';
  @Output() select = new EventEmitter<any>()

  constructor() { }
  ngOnInit(): void { }

}
