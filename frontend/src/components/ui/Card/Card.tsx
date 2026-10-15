import type { HTMLAttributes, ReactNode } from 'react'
import './Card.css'

export interface CardProps extends HTMLAttributes<HTMLDivElement> {
  variant?: 'default' | 'outlined' | 'elevated'
  padding?: 'none' | 'sm' | 'md' | 'lg'
  hoverable?: boolean
  children: ReactNode
}

export const Card = ({
  variant = 'default',
  padding = 'md',
  hoverable = false,
  children,
  className = '',
  ...props
}: CardProps) => {
  const classNames = [
    'card',
    `card-${variant}`,
    `card-padding-${padding}`,
    hoverable ? 'card-hoverable' : '',
    className,
  ]
    .filter(Boolean)
    .join(' ')

  return <div className={classNames} {...props}>{children}</div>
}

export interface CardHeaderProps extends HTMLAttributes<HTMLDivElement> {
  icon?: ReactNode
  title: string
  subtitle?: ReactNode
  action?: ReactNode
}

export const CardHeader = ({ icon, title, subtitle, action, className = '', ...props }: CardHeaderProps) => {
  return (
    <div className={`card-header ${className}`} {...props}>
      {icon && <div className="card-header-icon">{icon}</div>}
      <div className="card-header-content">
        <h3 className="card-header-title">{title}</h3>
        {subtitle && <p className="card-header-subtitle">{subtitle}</p>}
      </div>
      {action && <div className="card-header-action">{action}</div>}
    </div>
  )
}

export interface CardBodyProps extends HTMLAttributes<HTMLDivElement> {}

export const CardBody = ({ children, className = '', ...props }: CardBodyProps) => {
  return <div className={`card-body ${className}`} {...props}>{children}</div>
}

export interface CardFooterProps extends HTMLAttributes<HTMLDivElement> {}

export const CardFooter = ({ children, className = '', ...props }: CardFooterProps) => {
  return <div className={`card-footer ${className}`} {...props}>{children}</div>
}