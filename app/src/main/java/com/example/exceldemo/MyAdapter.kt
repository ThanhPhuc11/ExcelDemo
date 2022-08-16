package com.example.exceldemo

import android.annotation.SuppressLint
import android.view.LayoutInflater
import android.view.ViewGroup
import androidx.recyclerview.widget.RecyclerView
import com.example.exceldemo.databinding.ItemUserBinding

class MyAdapter(private val mList: List<UserModel>) :
    RecyclerView.Adapter<MyAdapter.MyViewHolder>() {
    var onClickItem: ((Int) -> Unit)? = null
    inner class MyViewHolder(val itemBinding: ItemUserBinding) :
        RecyclerView.ViewHolder(itemBinding.root)

    override fun onCreateViewHolder(parent: ViewGroup, viewType: Int): MyViewHolder {
        val itemBinding =
            ItemUserBinding.inflate(
                LayoutInflater.from(parent.context),
                parent,
                false
            )
        return MyViewHolder(itemBinding)
    }

    @SuppressLint("SetTextI18n")
    override fun onBindViewHolder(holder: MyViewHolder, position: Int) {
        val obj = mList[position]
        holder.itemBinding.tvNo.text = (position + 1).toString()
        holder.itemBinding.tvName.text = obj.name

        holder.itemView.setOnClickListener {
            onClickItem?.invoke(position)
        }
    }

    override fun getItemCount(): Int {
        return mList.size
    }
}